// XLSX 网格编辑器：xlsxToJson → sheets/rows/cells 模型 ↔ DOM 双向同步 → jsonToXlsx 往返。
// A1 风格网格：列头字母 + 行号，单元格可编辑（支持公式 =），工具栏作用于选区，
// 支持多表切换、增删行列、合并单元格、字体/对齐/填充样式。
import { xlsxToJson, jsonToXlsx } from '@fefeding/jdoc';
import type {
    XlsxCellObject,
    XlsxCellStyle,
    XlsxFontStyle,
    XlsxMetadata,
    XlsxRow,
    XlsxSheet,
    XlsxWorkbook
} from '@fefeding/jdoc';
import { EditorBase, type EditorHooks } from './framework';
import { h, colName } from './util';
import { icon } from './icons';

interface Model {
    metadata?: XlsxMetadata;
    sheets: XlsxSheet[];
    assets?: unknown[];
}

interface Range {
    r1: number;
    c1: number;
    r2: number;
    c2: number;
}

/** 'B3' → { c: 1, r: 2 }（0-based） */
function fromA1(s: string): { r: number; c: number } {
    const m = s.match(/^([A-Z]+)(\d+)$/i);
    if (!m) return { r: 0, c: 0 };
    let c = 0;
    for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: parseInt(m[2], 10) - 1, c: c - 1 };
}

/** { r, c }（0-based）→ 'B3' */
function toA1(r: number, c: number): string {
    return colName(c + 1) + (r + 1);
}

/** Excel 日期序列号(1899-12-30 起的天数) → Date（本地时区近似） */
function serialToDate(serial: number): Date | null {
    if (!isFinite(serial) || serial < 0) return null;
    return new Date(Math.round((serial - 25569) * 86400 * 1000));
}
function pad(n: number, w = 2): string {
    return String(Math.abs(n)).padStart(w, '0');
}
/** 按 numFmt 格式化数值显示（百分比/千分位/小数位/货币/日期）。 */
function formatNum(value: unknown, fmt: string): string {
    if (!fmt || fmt === 'General' || fmt === '常规') return value == null ? '' : String(value);
    const n = typeof value === 'number' ? value : parseFloat(String(value));
    if (isNaN(n)) return String(value ?? '');
    if (/yyyy|yy|mm|dd|hh|ss/i.test(fmt)) {
        const d = serialToDate(n);
        if (d) {
            return fmt
                .replace(/yyyy/i, String(d.getFullYear()))
                .replace(/yy/i, pad(d.getFullYear() % 100))
                .replace(/mm/i, pad(d.getMonth() + 1))
                .replace(/dd/i, pad(d.getDate()))
                .replace(/hh/i, pad(d.getHours()))
                .replace(/ss/i, pad(d.getSeconds()));
        }
        return String(value);
    }
    const pct = fmt.includes('%');
    const v = pct ? n * 100 : n;
    const decMatch = fmt.match(/\.(0+)/);
    const dec = decMatch ? decMatch[1].length : 0;
    const thousands = fmt.includes(',');
    let s = Math.abs(v).toFixed(dec);
    if (thousands) s = s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const cur = fmt.match(/[¥$€£]/);
    if (cur) s = cur[0] + s;
    s = (v < 0 ? '-' : '') + s;
    return pct ? s + '%' : s;
}

function cellDisplay(cell: XlsxCellObject | null): string {
    if (!cell) return '';
    let raw: string;
    if (cell.formula) raw = cell.result != null ? String(cell.result) : '=' + cell.formula;
    else if (cell.text != null) raw = cell.text;
    else if (cell.richText && cell.richText.length) raw = cell.richText.map((r) => r.text || '').join('');
    else if (cell.value != null) raw = String(cell.value);
    else raw = '';
    const fmt = cell.style?.numFmt;
    if (fmt) {
        if (!cell.formula && cell.value != null && typeof cell.value !== 'string') raw = formatNum(cell.value, fmt);
        else if (cell.formula && cell.result != null && typeof cell.result !== 'string') raw = formatNum(cell.result, fmt);
    }
    return raw;
}

function ensureCell(sheet: XlsxSheet, r: number, c: number): XlsxCellObject {
    const row = sheet.rows[r] as XlsxRow;
    if (!row.cells) (row as { cells: (XlsxCellObject | null)[] }).cells = [];
    let cell = row.cells[c];
    if (!cell) {
        cell = {};
        row.cells[c] = cell;
    }
    return cell;
}

function fontOf(cell: XlsxCellObject | null): XlsxFontStyle {
    if (!cell) return {};
    if (!cell.style) cell.style = {};
    if (!cell.style.font) cell.style.font = {};
    return cell.style.font;
}
function styleOf(cell: XlsxCellObject | null): XlsxCellStyle {
    if (!cell) return {};
    if (!cell.style) cell.style = {};
    return cell.style;
}

export class XlsxEditor extends EditorBase {
    private model: Model = { sheets: [] };
    private activeSheet = 0;
    private sel: Range = { r1: 0, c1: 0, r2: 0, c2: 0 };
    private gridWrap: HTMLElement | null = null;
    private tabsEl: HTMLElement | null = null;
    private editing: { r: number; c: number } | null = null;
    private statusRef?: (s: string) => void;
    private past: XlsxSheet[][] = [];
    private future: XlsxSheet[][] = [];
    private undoBtn: HTMLButtonElement | null = null;
    private redoBtn: HTMLButtonElement | null = null;

    async load(bytes: Uint8Array): Promise<void> {
        const parsed = await xlsxToJson(bytes, { assets: 'inline' });
        this.model = {
            metadata: parsed.metadata,
            sheets: (parsed.sheets || []) as XlsxSheet[],
            assets: parsed.assets as unknown[] | undefined
        };
        this.activeSheet = 0;
        this.recomputeAllSheets();
        this.buildToolbar();
        this.renderTabs();
        this.renderGrid();
    }

    protected async serialize(): Promise<Uint8Array> {
        // 深拷贝快照：避免 jsonToXlsx(async) 序列化期间模型被再次修改而产出中间态。
        const wb: XlsxWorkbook = { metadata: this.model.metadata, sheets: this.cloneSheets() };
        return jsonToXlsx(wb, this.model.assets ? { assets: this.model.assets as never } : {});
    }

    destroy(): void {
        this.gridWrap = null;
        this.tabsEl = null;
    }

    // ---------- 撤销 / 重做 ----------
    private cloneSheets(): XlsxSheet[][] {
        return JSON.parse(JSON.stringify(this.model.sheets)) as XlsxSheet[][];
    }
    private checkpoint(): void {
        this.past.push(this.cloneSheets());
        if (this.past.length > 200) this.past.shift();
        this.future = [];
        this.updateUndoRedo();
    }
    undo(): void {
        if (!this.past.length) return;
        this.future.push(this.cloneSheets());
        this.model.sheets = this.past.pop()!;
        if (this.activeSheet >= this.model.sheets.length) this.activeSheet = this.model.sheets.length - 1;
        this.renderTabs();
        this.paintGrid();
        this.updateUndoRedo();
        if (this.past.length === 0) this.syncRevert();
        else this.markDirty();
    }
    redo(): void {
        if (!this.future.length) return;
        this.past.push(this.cloneSheets());
        this.model.sheets = this.future.pop()!;
        if (this.activeSheet >= this.model.sheets.length) this.activeSheet = this.model.sheets.length - 1;
        this.renderTabs();
        this.paintGrid();
        this.updateUndoRedo();
        if (this.past.length === 0) this.syncRevert();
        else this.markDirty();
    }
    private updateUndoRedo(): void {
        if (this.undoBtn) this.undoBtn.disabled = !this.past.length;
        if (this.redoBtn) this.redoBtn.disabled = !this.future.length;
    }

    private sheet(): XlsxSheet | undefined {
        return this.model.sheets[this.activeSheet];
    }

    // ---------- 工具栏 ----------
    private buildToolbar(): void {
        const bar = this.toolbarEl;
        if (!bar) return;
        bar.innerHTML = '';
        this.undoBtn = h('button', { class: 'oe-tb', title: '撤销 (Ctrl+Z)', html: icon('undo'), onmousedown: (e: Event) => e.preventDefault(), onclick: () => this.undo() });
        this.redoBtn = h('button', { class: 'oe-tb', title: '重做 (Ctrl+Shift+Z)', html: icon('redo'), onmousedown: (e: Event) => e.preventDefault(), onclick: () => this.redo() });
        bar.append(this.undoBtn, this.redoBtn, h('span', { class: 'oe-sep' }));
        const tb = (ic: string, title: string, fn: () => void, glyph = '') =>
            h('button', { class: 'oe-tb' + (glyph ? ' oe-glyph' : ''), title, html: glyph || icon(ic), onclick: fn });
        const sizeSel = h('select', { class: 'oe-select', title: '字号', style: 'width:60px', onchange: (e: Event) => this.setFontSize(parseInt((e.target as HTMLSelectElement).value, 10)) });
        for (const p of [8, 9, 10, 11, 12, 14, 16, 18, 24, 36]) sizeSel.append(h('option', { value: String(p), text: String(p) }));
        const fmtSel = h('select', { class: 'oe-select', title: '数字格式', style: 'width:104px', onchange: (e: Event) => this.setNumFmt((e.target as HTMLSelectElement).value) });
        for (const [v, l] of [['', '常规'], ['0', '整数'], ['0.00', '两位小数'], ['0.00%', '百分比'], ['#,##0', '千分位'], ['#,##0.00', '千分位.两位'], ['¥#,##0.00', '货币'], ['yyyy-mm-dd', '日期']] as [string, string][]) fmtSel.append(h('option', { value: v, text: l }));
        bar.append(
            tb('bold', '加粗', () => this.fontToggle('bold'), 'B'),
            tb('italic', '斜体', () => this.fontToggle('italic'), 'I'),
            tb('underline', '下划线', () => this.fontToggle('underline'), 'U'),
            tb('strike', '删除线', () => this.fontToggle('strike'), 'S'),
            sizeSel,
            fmtSel,
            h('span', { class: 'oe-sep' }),
            tb('alignLeft', '左对齐', () => this.setAlign('left')),
            tb('alignCenter', '居中', () => this.setAlign('center')),
            tb('alignRight', '右对齐', () => this.setAlign('right')),
            h('span', { class: 'oe-sep' }),
            h('input', { type: 'color', class: 'oe-color', title: '填充色', value: '#ffffff', oninput: (e: Event) => this.setFill((e.target as HTMLInputElement).value) }),
            h('input', { type: 'color', class: 'oe-color', title: '文字色', value: '#000000', oninput: (e: Event) => this.setFontColor((e.target as HTMLInputElement).value) }),
            h('span', { class: 'oe-sep' }),
            tb('plus', '在上方插入行', () => this.insertRow('above')),
            tb('down', '在下方插入行', () => this.insertRow('below')),
            tb('trash', '删除行', () => this.deleteRow()),
            tb('plus', '在左侧插入列', () => this.insertCol('left')),
            tb('down', '在右侧插入列', () => this.insertCol('right')),
            tb('trash', '删除列', () => this.deleteCol()),
            h('span', { class: 'oe-sep' }),
            tb('table', '合并单元格', () => this.mergeCells()),
            h('span', { class: 'oe-sep' }),
            tb('table', '冻结到选区(上方行+左侧列)', () => this.setFreeze(this.sel.r1, this.sel.c1)),
            tb('table', '冻结首行', () => this.setFreeze(1, (this.sheet()?.frozen?.columns) || 0)),
            tb('table', '冻结首列', () => this.setFreeze((this.sheet()?.frozen?.rows) || 0, 1)),
            tb('table', '取消冻结', () => this.setFreeze(0, 0)),
            h('span', { class: 'oe-spacer' })
        );
    }

    // ---------- 渲染 ----------
    private renderTabs(): void {
        const host = this.tabsEl;
        if (!host) return;
        host.innerHTML = '';
        this.model.sheets.forEach((s, i) => {
            const btn = h('button', { class: i === this.activeSheet ? 'active' : '', text: s.name, onclick: () => { this.activeSheet = i; this.sel = { r1: 0, c1: 0, r2: 0, c2: 0 }; this.renderTabs(); this.renderGrid(); } });
            host.append(btn);
        });
        host.append(h('button', { class: 'add', title: '新建工作表', text: '+', onclick: () => this.addSheet() }));
    }

    private renderGrid(): void {
        const content = this.contentEl;
        if (!content) return;
        content.innerHTML = '';
        const wrap = h('div', { class: 'oe-xl' });
        const tabs = h('div', { class: 'oe-xl-tabs' });
        this.tabsEl = tabs;
        this.renderTabs();
        const gridWrap = h('div', { class: 'oe-grid-wrap' });
        this.gridWrap = gridWrap;
        wrap.append(tabs, gridWrap);
        content.append(wrap);
        this.paintGrid();
    }

    private paintGrid(): void {
        const host = this.gridWrap;
        const sheet = this.sheet();
        if (!host || !sheet) return;
        host.innerHTML = '';
        const rows = (sheet.rows || []) as XlsxRow[];
        const colCount = Math.max(1, ...rows.map((r) => (r.cells || []).length), (sheet.columns || []).length);
        const merges = this.mergesOf(sheet);
        const covered = new Set<string>(); // 被合并覆盖的 cell（跳过渲染）
        const anchors = new Map<string, { rs: number; cs: number }>(); // 左上角 → span
        for (const m of merges) {
            anchors.set(m.r1 + ':' + m.c1, { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
            for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (!(r === m.r1 && c === m.c1)) covered.add(r + ':' + c);
        }

        // 冻结窗格与列宽/行高
        const fRows = (sheet.frozen && sheet.frozen.rows) || 0;
        const fCols = (sheet.frozen && sheet.frozen.columns) || 0;
        const colWidths: number[] = [];
        for (let c = 0; c < colCount; c++) colWidths.push((sheet.columns && sheet.columns[c] && sheet.columns[c].width) || 96);
        const rowHeight = (i: number) => (rows[i] && rows[i].height) || 26;
        const colSpanW = (c: number, cs: number) => Array.from({ length: cs }, (_, k) => colWidths[c + k] || 96).reduce((a, b) => a + b, 0);

        const table = h('table', { class: 'oe-grid' });
        const colgroup = h('colgroup');
        colgroup.append(h('col', { style: 'width:48px' }));
        for (let c = 0; c < colCount; c++) colgroup.append(h('col', { style: `width:${colWidths[c]}px` }));
        table.append(colgroup);

        const thead = h('thead');
        const htr = h('tr');
        htr.append(h('th', { class: 'oe-corner' }));
        for (let c = 0; c < colCount; c++) {
            const th = h('th', { text: colName(c + 1), title: '列 ' + colName(c + 1) + '（拖拽边缘调整宽度）' });
            th.append(h('span', { class: 'col-res', dataset: { c: String(c) }, onmousedown: (e: MouseEvent) => { e.preventDefault(); e.stopPropagation(); this.colResize(c, e); } }));
            htr.append(th);
        }
        thead.append(htr);
        table.append(thead);

        const tbody = h('tbody');
        rows.forEach((row, r) => {
            const tr = h('tr');
            if (row.height) tr.style.height = row.height + 'pt';
            const frozenRow = r < fRows;
            const topPx = 26 + Array.from({ length: r }, (_, k) => rowHeight(k)).reduce((a, b) => a + b, 0);
            const rn = h('th', { class: 'oe-rownum' + (frozenRow ? ' fz' : ''), text: String(r + 1) });
            if (frozenRow) rn.style.top = topPx + 'px';
            tr.append(rn);
            let leftPx = 48;
            for (let c = 0; c < colCount; c++) {
                if (covered.has(r + ':' + c)) {
                    leftPx += colWidths[c];
                    continue; // 被合并覆盖：占位但不渲染 td
                }
                const td = h('td', { dataset: { r: String(r), c: String(c) } });
                const span = anchors.get(r + ':' + c);
                if (span) {
                    if (span.rs > 1) td.rowSpan = span.rs;
                    if (span.cs > 1) td.colSpan = span.cs;
                }
                if (frozenRow) {
                    td.classList.add('fz');
                    td.style.top = topPx + 'px';
                }
                if (c < fCols) {
                    td.classList.add('fz');
                    td.style.left = leftPx + 'px';
                }
                leftPx += span && span.cs > 1 ? colSpanW(c, span.cs) : colWidths[c];
                this.paintCell(td, sheet, r, c);
                tr.append(td);
            }
            tbody.append(tr);
        });
        table.append(tbody);
        host.append(table);
        this.updateSelectionVisual();
        this.updateUndoRedo();
    }

    private paintCell(td: HTMLElement, sheet: XlsxSheet, r: number, c: number): void {
        const row = sheet.rows[r] as XlsxRow;
        const cell = row?.cells ? row.cells[c] : null;
        const isEditing = this.editing && this.editing.r === r && this.editing.c === c;
        const el = h('div', { class: 'cell', text: cellDisplay(cell || null) });
        const st = cell?.style;
        if (st) {
            if (st.font?.bold) el.style.fontWeight = '700';
            if (st.font?.italic) el.style.fontStyle = 'italic';
            const deco: string[] = [];
            if (st.font?.underline) deco.push('underline');
            if (st.font?.strike) deco.push('line-through');
            if (deco.length) el.style.textDecoration = deco.join(' ');
            if (st.font?.color) el.style.color = st.font.color;
            if (st.font?.size) el.style.fontSize = st.font.size + 'pt';
            if (st.align) el.style.textAlign = st.align;
            if (st.fill) el.style.background = st.fill;
        }
        if (isEditing) {
            el.contentEditable = 'true';
            el.textContent = cell?.formula ? '=' + cell.formula : cellDisplay(cell || null);
        }
        td.append(el);
        td.addEventListener('mousedown', (e: MouseEvent) => this.onCellMouseDown(e, r, c));
        td.addEventListener('dblclick', () => this.beginEdit(r, c));
    }

    // ---------- 选区 / 编辑 ----------
    private onCellMouseDown(e: MouseEvent, r: number, c: number): void {
        this.commitEdit();
        if (e.shiftKey) {
            this.sel.r2 = r;
            this.sel.c2 = c;
        } else {
            this.sel = { r1: r, c1: c, r2: r, c2: c };
        }
        this.normalizeSel();
        this.updateSelectionVisual();
    }

    private normalizeSel(): void {
        const s = this.sel;
        s.r1 = Math.min(s.r1, s.r2);
        s.c1 = Math.min(s.c1, s.c2);
        s.r2 = Math.max(s.r1, s.r2);
        s.c2 = Math.max(s.c1, s.c2);
    }

    private inSel(r: number, c: number): boolean {
        return r >= this.sel.r1 && r <= this.sel.r2 && c >= this.sel.c1 && c <= this.sel.c2;
    }

    private updateSelectionVisual(): void {
        this.gridWrap?.querySelectorAll('td').forEach((td) => {
            const r = Number((td as HTMLElement).dataset.r);
            const c = Number((td as HTMLElement).dataset.c);
            td.classList.toggle('inrange', this.inSel(r, c));
            td.classList.toggle('active', r === this.sel.r1 && c === this.sel.c1);
        });
        this.updateStatus();
    }

    private beginEdit(r: number, c: number): void {
        this.checkpoint();
        this.editing = { r, c };
        this.paintGrid();
        const td = this.gridWrap?.querySelector(`td[data-r="${r}"][data-c="${c}"]`) as HTMLElement | null;
        const el = td?.querySelector('.cell') as HTMLElement | null;
        if (el) {
            el.focus();
            const range = document.createRange();
            range.selectNodeContents(el);
            const selObj = window.getSelection();
            selObj?.removeAllRanges();
            selObj?.addRange(range);
            el.addEventListener('keydown', (ev: KeyboardEvent) => this.onEditKey(ev, r, c));
            el.addEventListener('blur', () => this.commitEdit());
        }
    }

    private onEditKey(e: KeyboardEvent, r: number, c: number): void {
        if (e.key === 'Enter') {
            e.preventDefault();
            this.commitEdit();
            this.moveActive(1, 0);
        } else if (e.key === 'Tab') {
            e.preventDefault();
            this.commitEdit();
            this.moveActive(0, e.shiftKey ? -1 : 1);
        } else if (e.key === 'Escape') {
            this.editing = null;
            this.paintGrid();
        }
    }

    private moveActive(dr: number, dc: number): void {
        const sheet = this.sheet();
        if (!sheet) return;
        const maxR = (sheet.rows || []).length - 1;
        const maxC = Math.max(0, ...((sheet.rows || []) as XlsxRow[]).map((rw) => (rw.cells || []).length)) - 1;
        const r = Math.max(0, Math.min(maxR, this.sel.r1 + dr));
        const c = Math.max(0, Math.min(maxC, this.sel.c1 + dc));
        this.sel = { r1: r, c1: c, r2: r, c2: c };
        this.updateSelectionVisual();
    }

    private commitEdit(): void {
        if (!this.editing) return;
        const { r, c } = this.editing;
        const sheet = this.sheet();
        const td = this.gridWrap?.querySelector(`td[data-r="${r}"][data-c="${c}"]`) as HTMLElement | null;
        const el = td?.querySelector('.cell') as HTMLElement | null;
        this.editing = null;
        if (sheet && el) this.writeCell(sheet, r, c, el.innerText || '');
        this.paintGrid();
    }

    private writeCell(sheet: XlsxSheet, r: number, c: number, raw: string): void {
        const input = raw.trim();
        const cell = ensureCell(sheet, r, c);
        delete cell.value;
        delete cell.text;
        delete cell.formula;
        delete cell.result;
        if (input === '') {
            // 清空
        } else if (input.startsWith('=')) {
            cell.formula = input.slice(1);
            try {
                cell.result = this.evaluateFormula(input.slice(1), sheet);
            } catch {
                cell.result = null;
            }
        } else if (input !== '' && !isNaN(Number(input)) && /^-?\d*(\.\d+)?$/.test(input)) {
            cell.value = Number(input);
        } else {
            cell.value = input;
        }
        // 无论改的是值还是公式，都沿依赖图增量重算受影响的公式（含跨表）
        this.recomputeDependencies([{ sheet: sheet.name, r, c }]);
        this.markDirty();
    }

    // ---------- 样式命令 ----------
    private eachSelected(fn: (cell: XlsxCellObject, r: number, c: number) => void): void {
        const sheet = this.sheet();
        if (!sheet) return;
        this.checkpoint();
        for (let r = this.sel.r1; r <= this.sel.r2; r++) {
            for (let c = this.sel.c1; c <= this.sel.c2; c++) {
                if (!(sheet.rows[r] as XlsxRow | undefined)) continue;
                fn(ensureCell(sheet, r, c), r, c);
            }
        }
    }

    private fontToggle(key: keyof XlsxFontStyle): void {
        const active = this.selActiveCell();
        const on = !(active && active.style?.font && (active.style.font as Record<string, unknown>)[key]);
        this.eachSelected((cell) => {
            fontOf(cell)[key] = on as never;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setFontSize(size: number): void {
        this.eachSelected((cell) => {
            fontOf(cell).size = size;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setNumFmt(fmt: string): void {
        this.eachSelected((cell) => {
            if (!cell.style) cell.style = {};
            if (fmt) cell.style.numFmt = fmt;
            else delete cell.style.numFmt;
        });
        this.paintGrid();
        this.markDirty();
    }

    private selActiveCell(): XlsxCellObject | null {
        const sheet = this.sheet();
        const row = sheet?.rows[this.sel.r1] as XlsxRow | undefined;
        return row?.cells ? row.cells[this.sel.c1] : null;
    }

    private setAlign(a: 'left' | 'center' | 'right'): void {
        this.eachSelected((cell) => {
            styleOf(cell).align = a;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setFill(color: string): void {
        this.eachSelected((cell) => {
            styleOf(cell).fill = color;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setFontColor(color: string): void {
        this.eachSelected((cell) => {
            fontOf(cell).color = color;
        });
        this.paintGrid();
        this.markDirty();
    }

    // ---------- 行 / 列 ----------
    private reindex(sheet: XlsxSheet): void {
        (sheet.rows as XlsxRow[]).forEach((row, i) => {
            row.index = i + 1;
        });
    }
    private insertRow(where: 'above' | 'below'): void {
        const sheet = this.sheet();
        if (!sheet) return;
        this.checkpoint();
        const at = where === 'above' ? this.sel.r1 : this.sel.r2 + 1;
        const colCount = Math.max(1, ...((sheet.rows || []) as XlsxRow[]).map((r) => (r.cells || []).length));
        (sheet.rows as XlsxRow[]).splice(at, 0, { index: at + 1, cells: new Array(colCount).fill(null) });
        this.reindex(sheet);
        this.paintGrid();
        this.markDirty();
    }
    private deleteRow(): void {
        const sheet = this.sheet();
        if (!sheet || (sheet.rows || []).length <= 1) return;
        this.checkpoint();
        const count = this.sel.r2 - this.sel.r1 + 1;
        (sheet.rows as XlsxRow[]).splice(this.sel.r1, count);
        this.reindex(sheet);
        this.sel = { r1: Math.min(this.sel.r1, (sheet.rows as XlsxRow[]).length - 1), c1: this.sel.c1, r2: Math.min(this.sel.r2, (sheet.rows as XlsxRow[]).length - 1), c2: this.sel.c2 };
        this.normalizeSel();
        this.paintGrid();
        this.markDirty();
    }
    private insertCol(where: 'left' | 'right'): void {
        const sheet = this.sheet();
        if (!sheet) return;
        this.checkpoint();
        const at = where === 'left' ? this.sel.c1 : this.sel.c2 + 1;
        for (const row of sheet.rows as XlsxRow[]) {
            if (!row.cells) (row as { cells: (XlsxCellObject | null)[] }).cells = [];
            row.cells.splice(at, 0, null);
        }
        this.paintGrid();
        this.markDirty();
    }
    private deleteCol(): void {
        const sheet = this.sheet();
        if (!sheet) return;
        this.checkpoint();
        const count = this.sel.c2 - this.sel.c1 + 1;
        for (const row of sheet.rows as XlsxRow[]) row.cells?.splice(this.sel.c1, count);
        this.paintGrid();
        this.markDirty();
    }

    // ---------- 列宽拖拽 ----------
    private colResize(c: number, e: MouseEvent): void {
        const sheet = this.sheet();
        const host = this.gridWrap;
        if (!sheet || !host) return;
        const cols = host.querySelectorAll('colgroup col');
        const targetCol = cols[c + 1] as HTMLElement | undefined;
        const w0 = (sheet.columns && sheet.columns[c] && sheet.columns[c].width) || 96;
        const startX = e.clientX;
        const onMove = (ev: MouseEvent) => {
            const w = Math.max(24, w0 + ev.clientX - startX);
            if (targetCol) targetCol.style.width = w + 'px';
        };
        const onUp = (ev: MouseEvent) => {
            const w = Math.max(24, Math.round(w0 + ev.clientX - startX));
            this.checkpoint();
            if (!sheet.columns) sheet.columns = [];
            if (!sheet.columns[c]) sheet.columns[c] = {};
            sheet.columns[c].width = w;
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            this.markDirty();
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
    }

    private setFreeze(rows: number, columns: number): void {
        const sheet = this.sheet();
        if (!sheet) return;
        this.checkpoint();
        sheet.frozen = rows || columns ? { rows, columns } : undefined;
        this.paintGrid();
        this.markDirty();
    }

    // ---------- 轻量公式求值 ----------
    /** 重算一张表内所有公式单元格的 result（按行序两遍以尽量满足依赖）。 */
    private recomputeSheet(sheet: XlsxSheet): void {
        for (let pass = 0; pass < 2; pass++) {
            for (const row of sheet.rows as XlsxRow[]) {
                for (const cell of row.cells || []) {
                    if (cell && cell.formula != null) {
                        try {
                            cell.result = this.evaluateFormula(cell.formula, sheet);
                        } catch {
                            cell.result = null;
                        }
                    }
                }
            }
        }
    }

    /** 重算所有工作表的公式（两遍），支持跨表引用在全部表就绪后求值。 */
    private recomputeAllSheets(): void {
        for (let pass = 0; pass < 2; pass++) for (const s of this.model.sheets) this.recomputeSheet(s);
    }

    /** 提取公式中的单元格/区域引用（含跨表 Sheet!A1 / 'Sheet'!A1:B2）。 */
    private extractRefs(formula: string): { sheet?: string; ref: string }[] {
        const src = String(formula);
        const out: { sheet?: string; ref: string }[] = [];
        const cellRe = /^[A-Za-z]{1,3}\d+/;
        let i = 0;
        while (i < src.length) {
            const ch = src[i];
            if (ch === '"' || ch === "'") { const e = src.indexOf(ch, i + 1); i = e < 0 ? src.length : e + 1; continue; }
            if (/[A-Za-z]/.test(ch)) {
                const word = src.slice(i).match(/^[A-Za-z]+[A-Za-z0-9]*/)![0];
                if (src[i + word.length] === '(') { i += word.length; continue; }
                const at = src[i + word.length] === '!' ? i + word.length + 1 : -1;
                if (at >= 0) {
                    const rm = src.slice(at).match(cellRe);
                    if (rm) {
                        let ref = rm[0];
                        let j = at + rm[0].length;
                        if (src[j] === ':') { const rm2 = src.slice(j + 1).match(cellRe); if (rm2) { ref += ':' + rm2[0]; j += 1 + rm2[0].length; } }
                        out.push({ sheet: word, ref: ref.toUpperCase() });
                        i = j;
                        continue;
                    }
                }
                if (cellRe.test(word)) {
                    let ref = word;
                    let j = i + word.length;
                    if (src[j] === ':') { const rm2 = src.slice(j + 1).match(cellRe); if (rm2) { ref += ':' + rm2[0]; j += 1 + rm2[0].length; } }
                    out.push({ ref: ref.toUpperCase() });
                    i = j;
                    continue;
                }
                i += word.length;
                continue;
            }
            i++;
        }
        return out;
    }

    /** 判断 (r,c)（0-based）是否落在引用（单格或区域）内。 */
    private refCovers(refStr: string, r: number, c: number): boolean {
        const parts = refStr.split(':');
        const a = fromA1(parts[0]);
        const b = parts[1] ? fromA1(parts[1]) : a;
        return r >= Math.min(a.r, b.r) && r <= Math.max(a.r, b.r) && c >= Math.min(a.c, b.c) && c <= Math.max(a.c, b.c);
    }

    /** 依赖图增量重算：从变更单元格出发，BFS 重算所有（跨表）依赖它的公式，直到稳定。 */
    private recomputeDependencies(changed: { sheet: string; r: number; c: number }[]): void {
        const formulas: { sheet: XlsxSheet; cell: XlsxCellObject; r: number; c: number }[] = [];
        for (const s of this.model.sheets)
            (s.rows as XlsxRow[]).forEach((row, r) => (row.cells || []).forEach((cell, c) => { if (cell && cell.formula != null) formulas.push({ sheet: s, cell, r, c }); }));
        let frontier = changed.slice();
        const visited = new Set<string>();
        let guard = 0;
        while (frontier.length && guard++ < 100) {
            const next: { sheet: string; r: number; c: number }[] = [];
            for (const ch of frontier) {
                for (const f of formulas) {
                    const key = f.sheet.name + '!' + f.r + ':' + f.c;
                    if (visited.has(key)) continue;
                    const refs = this.extractRefs(f.cell.formula!);
                    const hit = refs.some((rf) => (rf.sheet || f.sheet.name) === ch.sheet && this.refCovers(rf.ref, ch.r, ch.c));
                    if (!hit) continue;
                    visited.add(key);
                    const old = f.cell.result;
                    try {
                        f.cell.result = this.evaluateFormula(f.cell.formula!, f.sheet);
                    } catch {
                        f.cell.result = null;
                    }
                    if (old !== f.cell.result) next.push({ sheet: f.sheet.name, r: f.r, c: f.c });
                }
            }
            frontier = next;
        }
    }

    /** 求值公式（不含前导 =）：支持 + - * / ^ %、括号、单元格/区域引用、比较与常用函数。 */
    private evaluateFormula(formula: string, sheet: XlsxSheet): number | string {
        const src = String(formula).replace(/^=/, '');
        const num = (v: unknown): number => {
            const n = typeof v === 'number' ? v : parseFloat(String(v));
            return isNaN(n) ? 0 : n;
        };
        type Tok = { t: 'num' | 'ref' | 'func' | 'str' | 'op'; v?: string; n?: number; sheet?: string };
        const toks: Tok[] = [];
        let i = 0;
        while (i < src.length) {
            const ch = src[i];
            if (ch === ' ') { i++; continue; }
            if (/[\d.]/.test(ch)) { const m = src.slice(i).match(/^\d*\.?\d+/)!; toks.push({ t: 'num', n: parseFloat(m[0]) }); i += m[0].length; continue; }
            if (ch === "'") {
                const end = src.indexOf("'", i + 1);
                const name = src.slice(i + 1, end < 0 ? src.length : end);
                i = end < 0 ? src.length : end + 1;
                if (src[i] === '!') {
                    i++;
                    const rm = src.slice(i).match(/^[A-Za-z]+\d+/);
                    if (rm) { toks.push({ t: 'ref', v: rm[0].toUpperCase(), sheet: name }); i += rm[0].length; }
                } else {
                    toks.push({ t: 'str', v: name });
                }
                continue;
            }
            if (/[A-Za-z]/.test(ch)) {
                const word = src.slice(i).match(/^[A-Za-z]+[A-Za-z0-9]*/)![0];
                if (src[i + word.length] === '(') { toks.push({ t: 'func', v: word.toUpperCase() }); i += word.length; continue; }
                if (src[i + word.length] === '!') {
                    i += word.length + 1;
                    const rm = src.slice(i).match(/^[A-Za-z]+\d+/);
                    if (rm) { toks.push({ t: 'ref', v: rm[0].toUpperCase(), sheet: word }); i += rm[0].length; }
                    continue;
                }
                toks.push({ t: 'ref', v: word.toUpperCase() }); i += word.length; continue;
            }
            if (ch === '"') { const end = src.indexOf('"', i + 1); toks.push({ t: 'str', v: src.slice(i + 1, end < 0 ? src.length : end) }); i = end < 0 ? src.length : end + 1; continue; }
            if ((ch === '<' || ch === '>') && src[i + 1] === '=') { toks.push({ t: 'op', v: ch + '=' }); i += 2; continue; }
            if (ch === '<' && src[i + 1] === '>') { toks.push({ t: 'op', v: '<>' }); i += 2; continue; }
            if ('+-*/^%(),:='.includes(ch)) { toks.push({ t: 'op', v: ch }); i++; continue; }
            i++;
        }
        let p = 0;
        const cur = (): Tok | undefined => toks[p];
        const eat = (v: string): boolean => { const t = cur(); if (t && t.t === 'op' && t.v === v) { p++; return true; } return false; };
        const refVal = (ref: string, sheetName?: string): number | string => {
            const target = sheetName ? this.model.sheets.find((s) => s.name === sheetName) : sheet;
            if (!target) return 0;
            const { r, c } = fromA1(ref);
            const row = target.rows[r] as XlsxRow | undefined;
            const cell = row?.cells ? row.cells[c] : null;
            if (!cell) return 0;
            if (cell.value != null) return typeof cell.value === 'number' ? cell.value : String(cell.value);
            if (cell.result != null) return cell.result as number;
            return 0;
        };
        const rangeVals = (a: string, b: string, sheetName?: string): number[] => {
            const pa = fromA1(a);
            const pb = fromA1(b);
            const out: number[] = [];
            for (let r = Math.min(pa.r, pb.r); r <= Math.max(pa.r, pb.r); r++)
                for (let c = Math.min(pa.c, pb.c); c <= Math.max(pa.c, pb.c); c++) out.push(num(refVal(colName(c + 1) + (r + 1), sheetName)));
            return out;
        };
        const matchCrit = (crit: unknown) => {
            const s = String(crit);
            const m = s.match(/^\s*(>=|<=|<>|>|<|=)?\s*(.*)$/);
            const op = m && m[1] ? m[1] : '=';
            const rest = m ? m[2].trim() : s;
            const rhs = parseFloat(rest);
            const numeric = rest !== '' && !isNaN(rhs);
            return (v: number): boolean => {
                if (numeric) {
                    if (op === '>') return v > rhs;
                    if (op === '<') return v < rhs;
                    if (op === '>=') return v >= rhs;
                    if (op === '<=') return v <= rhs;
                    if (op === '<>') return v !== rhs;
                    return v === rhs;
                }
                return String(v) === s.trim();
            };
        };
        const callFunc = (name: string, args: (number | string | number[])[]): number | string => {
            const nums: number[] = args.flatMap((x) => (Array.isArray(x) ? x : [num(x)]));
            const arr = (i: number): number[] => (Array.isArray(args[i]) ? (args[i] as number[]) : [num(args[i])]);
            switch (name) {
                case 'SUM': return nums.reduce((s, n) => s + n, 0);
                case 'AVERAGE': return nums.length ? nums.reduce((s, n) => s + n, 0) / nums.length : 0;
                case 'COUNT': return nums.length;
                case 'MAX': return nums.length ? Math.max(...nums) : 0;
                case 'MIN': return nums.length ? Math.min(...nums) : 0;
                case 'PRODUCT': return nums.length ? nums.reduce((s, n) => s * n, 1) : 0;
                case 'ABS': return Math.abs(num(args[0]));
                case 'SQRT': return Math.sqrt(num(args[0]));
                case 'INT': return Math.floor(num(args[0]));
                case 'ROUND': { const d = args.length > 1 ? num(args[1]) : 0; const f = Math.pow(10, d); return Math.round(num(args[0]) * f) / f; }
                case 'ROUNDUP': { const d = args.length > 1 ? num(args[1]) : 0; const f = Math.pow(10, d); return Math.ceil(num(args[0]) * f) / f; }
                case 'ROUNDDOWN': { const d = args.length > 1 ? num(args[1]) : 0; const f = Math.pow(10, d); return Math.floor(num(args[0]) * f) / f; }
                case 'MOD': { const b = num(args[1]); return b === 0 ? NaN : num(args[0]) % b; }
                case 'POWER': return Math.pow(num(args[0]), num(args[1]));
                case 'IF': return num(args[0]) !== 0 ? (args[1] ?? 0) : (args[2] ?? 0);
                case 'IFERROR': { const v = num(args[0]); return isNaN(v) ? (args[1] ?? 0) : v; }
                case 'AND': return args.every((x) => num(x) !== 0) ? 1 : 0;
                case 'OR': return args.some((x) => num(x) !== 0) ? 1 : 0;
                case 'NOT': return num(args[0]) === 0 ? 1 : 0;
                case 'COUNTIF': { const m = matchCrit(args[1]); return arr(0).filter(m).length; }
                case 'SUMIF': { const m = matchCrit(args[1]); const src = arr(0); const sum = args.length > 2 ? arr(2) : src; return src.reduce((s, v, i) => s + (m(v) ? (sum[i] ?? 0) : 0), 0); }
                case 'AVERAGEIF': { const m = matchCrit(args[1]); const hit = arr(0).filter(m); return hit.length ? hit.reduce((s, n) => s + n, 0) / hit.length : 0; }
                case 'LEN': return String(args[0] ?? '').length;
                case 'LEFT': return String(args[0] ?? '').slice(0, Math.max(0, num(args[1])));
                case 'RIGHT': { const s = String(args[0] ?? ''); const n = Math.max(0, num(args[1])); return n ? s.slice(-n) : ''; }
                case 'MID': { const s = String(args[0] ?? ''); const st = Math.max(1, num(args[1])); const ln = Math.max(0, num(args[2])); return s.slice(st - 1, st - 1 + ln); }
                case 'UPPER': return String(args[0] ?? '').toUpperCase();
                case 'LOWER': return String(args[0] ?? '').toLowerCase();
                case 'TRIM': return String(args[0] ?? '').trim();
                case 'CONCATENATE': return args.map((x) => (Array.isArray(x) ? x.join('') : String(x))).join('');
                case 'TODAY': return Math.floor(Date.now() / 86400000 + 25569);
                case 'NOW': return Date.now() / 86400000 + 25569;
                default: return 0;
            }
        };
        const expr = (): number | string | number[] => {
            let v: number | string | number[] = addsub();
            while (cur() && cur().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(cur().v!)) {
                const op = cur().v!; p++; const r = addsub();
                let res = false;
                if (op === '=') res = num(v) === num(r); else if (op === '<>') res = num(v) !== num(r);
                else if (op === '<') res = num(v) < num(r); else if (op === '>') res = num(v) > num(r);
                else if (op === '<=') res = num(v) <= num(r); else res = num(v) >= num(r);
                v = res ? 1 : 0;
            }
            return v;
        };
        const addsub = (): number | string | number[] => {
            let v = muldiv();
            while (cur() && cur().t === 'op' && (cur().v === '+' || cur().v === '-')) {
                const op = cur().v!; p++; const r = muldiv();
                if (op === '+') v = typeof v === 'string' || typeof r === 'string' ? String(v) + String(r) : num(v) + num(r);
                else v = num(v) - num(r);
            }
            return v;
        };
        const muldiv = (): number | string | number[] => {
            let v = power();
            while (cur() && cur().t === 'op' && (cur().v === '*' || cur().v === '/')) {
                const op = cur().v!; p++; const r = num(power());
                v = op === '*' ? num(v) * r : r === 0 ? NaN : num(v) / r;
            }
            return v;
        };
        const power = (): number | string | number[] => {
            let v = unary();
            while (cur() && cur().t === 'op' && cur().v === '^') { p++; v = Math.pow(num(v), num(unary())); }
            return v;
        };
        const unary = (): number | string | number[] => {
            const t = cur();
            if (t && t.t === 'op' && t.v === '-') { p++; return -num(unary()); }
            if (t && t.t === 'op' && t.v === '+') { p++; return unary(); }
            return primary();
        };
        const primary = (): number | string | number[] => {
            const t = cur();
            if (!t) return 0;
            if (t.t === 'num') { p++; return t.n!; }
            if (t.t === 'str') { p++; return t.v!; }
            if (t.t === 'func') {
                const name = t.v!; p++; eat('(');
                const args: (number | string | number[])[] = [];
                if (!(cur() && cur().t === 'op' && cur().v === ')')) { args.push(expr()); while (eat(',')) args.push(expr()); }
                eat(')');
                return callFunc(name, args);
            }
            if (t.t === 'ref') {
                p++;
                if (cur() && cur().t === 'op' && cur().v === ':') { p++; const t2 = cur(); if (t2 && t2.t === 'ref') { p++; return rangeVals(t.v!, t2.v!, t.sheet || t2.sheet); } }
                return refVal(t.v!, t.sheet);
            }
            if (t.t === 'op' && t.v === '(') { p++; const v = expr(); eat(')'); return v; }
            p++;
            return 0;
        };
        try {
            const r = expr();
            return Array.isArray(r) ? num(r[0]) : r;
        } catch {
            return 0;
        }
    }

    // ---------- 合并 ----------
    private mergesOf(sheet: XlsxSheet): Range[] {
        return (sheet.merges || []).map((s) => {
            const parts = s.split(':');
            const a = fromA1(parts[0]);
            const b = fromA1(parts[1] || parts[0]);
            return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
        });
    }
    private mergeCells(): void {
        const sheet = this.sheet();
        if (!sheet) return;
        if (this.sel.r1 === this.sel.r2 && this.sel.c1 === this.sel.c2) {
            this.toast('请先选中一个区域再合并');
            return;
        }
        this.checkpoint();
        const a1 = toA1(this.sel.r1, this.sel.c1) + ':' + toA1(this.sel.r2, this.sel.c2);
        if (!sheet.merges) sheet.merges = [];
        // 移除与新区域重叠的旧合并
        const nr = { r1: this.sel.r1, c1: this.sel.c1, r2: this.sel.r2, c2: this.sel.c2 };
        sheet.merges = sheet.merges.filter((s) => {
            const m = this.mergesOf({ merges: [s] } as XlsxSheet)[0];
            return !(m.r1 <= nr.r2 && m.r2 >= nr.r1 && m.c1 <= nr.c2 && m.c2 >= nr.c1);
        });
        sheet.merges.push(a1);
        this.paintGrid();
        this.markDirty();
    }

    // ---------- 工作表 ----------
    private addSheet(): void {
        this.checkpoint();
        const name = 'Sheet' + (this.model.sheets.length + 1);
        this.model.sheets.push({ name, rows: [{ index: 1, cells: [] }, { index: 2, cells: [] }] });
        this.activeSheet = this.model.sheets.length - 1;
        this.renderTabs();
        this.paintGrid();
        this.markDirty();
    }

    private updateStatus(): void {
        const sheet = this.sheet();
        const cell = this.selActiveCell();
        this.statusRef?.(`${toA1(this.sel.r1, this.sel.c1)}${this.sel.r1 !== this.sel.r2 || this.sel.c1 !== this.sel.c2 ? ':' + toA1(this.sel.r2, this.sel.c2) : ''}${sheet ? ' · ' + sheet.name : ''}${cell?.formula ? ' =' + cell.formula : ''}`);
    }

    setStatusHandler(fn: (s: string) => void): void {
        this.statusRef = fn;
    }
}

export function createXlsxEditor(hooks: EditorHooks): XlsxEditor {
    return new XlsxEditor(hooks);
}
