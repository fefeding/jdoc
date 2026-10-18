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

function cellDisplay(cell: XlsxCellObject | null): string {
    if (!cell) return '';
    if (cell.formula) return cell.result != null ? String(cell.result) : '=' + cell.formula;
    if (cell.text != null) return cell.text;
    if (cell.richText && cell.richText.length) return cell.richText.map((r) => r.text || '').join('');
    if (cell.value != null) return String(cell.value);
    return '';
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

    async load(bytes: Uint8Array): Promise<void> {
        const parsed = await xlsxToJson(bytes, { assets: 'inline' });
        this.model = {
            metadata: parsed.metadata,
            sheets: (parsed.sheets || []) as XlsxSheet[],
            assets: parsed.assets as unknown[] | undefined
        };
        this.activeSheet = 0;
        this.buildToolbar();
        this.renderTabs();
        this.renderGrid();
    }

    protected async serialize(): Promise<Uint8Array> {
        const wb: XlsxWorkbook = { metadata: this.model.metadata, sheets: this.model.sheets };
        return jsonToXlsx(wb, this.model.assets ? { assets: this.model.assets as never } : {});
    }

    destroy(): void {
        this.gridWrap = null;
        this.tabsEl = null;
    }

    private sheet(): XlsxSheet | undefined {
        return this.model.sheets[this.activeSheet];
    }

    // ---------- 工具栏 ----------
    private buildToolbar(): void {
        const bar = this.toolbarEl;
        if (!bar) return;
        bar.innerHTML = '';
        const tb = (ic: string, title: string, fn: () => void, glyph = '') =>
            h('button', { class: 'oe-tb' + (glyph ? ' oe-glyph' : ''), title, html: glyph || icon(ic), onclick: fn });
        bar.append(
            tb('bold', '加粗', () => this.fontToggle('bold'), 'B'),
            tb('italic', '斜体', () => this.fontToggle('italic'), 'I'),
            tb('underline', '下划线', () => this.fontToggle('underline'), 'U'),
            tb('strike', '删除线', () => this.fontToggle('strike'), 'S'),
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
        const prev = { ...cell };
        delete cell.value;
        delete cell.text;
        delete cell.formula;
        delete cell.result;
        if (input === '') {
            // 清空
        } else if (input.startsWith('=')) {
            cell.formula = input.slice(1);
            cell.result = prev.result ?? null;
        } else if (input !== '' && !isNaN(Number(input)) && /^-?\d*(\.\d+)?$/.test(input)) {
            cell.value = Number(input);
        } else {
            cell.value = input;
        }
        this.markDirty();
    }

    // ---------- 样式命令 ----------
    private eachSelected(fn: (cell: XlsxCellObject, r: number, c: number) => void): void {
        const sheet = this.sheet();
        if (!sheet) return;
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
        sheet.frozen = rows || columns ? { rows, columns } : undefined;
        this.paintGrid();
        this.markDirty();
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
