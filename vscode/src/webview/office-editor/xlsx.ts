// XLSX 网格编辑器：xlsxToJson → sheets/rows/cells 模型 ↔ DOM 双向同步 → jsonToXlsx 往返。
//
// 观感对齐 Excel / WPS：列标行 + 行号列、真实列宽行高、边框/填充/字体/对齐、
// 数字格式、合并单元格、冻结窗格、公式栏、底部工作表标签、状态栏求和统计，
// 以及浮动图表（柱/条/饼/折线/面积）——图表按锚点定位，数据随单元格实时联动。
//
// 单元格可编辑（支持公式 =），公式依赖跨表增量重算；工具栏作用于选区。
import {
    xlsxToJson,
    jsonToXlsx,
    evaluateFormula,
    extractFormulaRefs,
    parseRangeRef,
    formatByPattern,
    renderChartSvg,
    xlsxChartToSpec
} from '@fefeding/jdoc';
import type {
    FormulaContext,
    FormulaValue,
    XlsxCellObject,
    XlsxCellStyle,
    XlsxChart,
    XlsxChartSeriesData,
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

/** 编辑器选项 */
export interface XlsxEditorOptions {
    /** 只读（查看模式）：不显示工具栏，不可编辑 */
    readonly?: boolean;
}

/** 行号列宽（px） */
const ROW_HEADER_W = 46;
/** 列标行高（px） */
const HEAD_H = 22;
/** 默认列宽（Excel 8.43 字符 ≈ 64px） */
const DEFAULT_COL_W = 64;
/** 默认行高（15pt ≈ 20px） */
const DEFAULT_ROW_H = 20;
/** 1 像素 = 9525 EMU */
const EMU_PER_PX = 9525;

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

/** 数字转显示文本（General：最多 10 位小数） */
function plainNumber(n: number): string {
    if (!isFinite(n)) return String(n);
    return String(Number(n.toFixed(10)));
}

/** 按 numFmt 格式化数值 */
function formatNum(value: unknown, fmt?: string): string {
    if (value instanceof Date) {
        // 日期单元格：序列号无时区，按 UTC 取年月日再套 numFmt
        const serial = Math.round((value.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
        if (fmt && fmt !== 'General' && /[ymdhs]/i.test(fmt)) {
            return formatByPattern(serial, fmt) || value.toISOString().slice(0, 10);
        }
        return value.toISOString().slice(0, 10);
    }
    if (typeof value !== 'number') return value == null ? '' : String(value);
    if (!fmt || fmt === 'General' || fmt === '常规') return plainNumber(value);
    const text = formatByPattern(value, fmt);
    return text || plainNumber(value);
}

/** 单元格取值（公式取计算结果） */
function cellInfo(cell: XlsxCellObject | null | undefined): { value: unknown; text: string } {
    if (!cell) return { value: null, text: '' };
    if (cell.formula !== undefined) {
        const v = cell.result ?? null;
        return { value: v, text: v != null ? String(v) : '' };
    }
    if (cell.value !== undefined && cell.value !== null) return { value: cell.value, text: String(cell.value) };
    if (cell.richText && cell.richText.length) {
        const text = cell.richText.map((p) => p.text || '').join('');
        return { value: text, text };
    }
    if (cell.text !== undefined && cell.text !== null) return { value: cell.text, text: String(cell.text) };
    return { value: null, text: '' };
}

/** 单元格显示文本 */
function cellDisplay(cell: XlsxCellObject | null | undefined): string {
    if (!cell) return '';
    if (cell.formula !== undefined) {
        if (cell.result === undefined || cell.result === null) return '=' + cell.formula;
        return formatNum(cell.result, cell.style?.numFmt);
    }
    if (cell.value !== undefined && cell.value !== null) return formatNum(cell.value, cell.style?.numFmt);
    if (cell.richText && cell.richText.length) return cell.richText.map((p) => p.text || '').join('');
    if (cell.text !== undefined && cell.text !== null) return String(cell.text);
    return '';
}

/** 公式源码（公式栏显示） */
function cellFormula(cell: XlsxCellObject | null | undefined): string {
    if (!cell) return '';
    if (cell.formula !== undefined) return '=' + cell.formula;
    const info = cellInfo(cell);
    return info.text;
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

/** 边框线型 -> CSS */
function borderCss(style: string | undefined, color: string | undefined): string {
    const c = color || '#BFBFBF';
    switch (style) {
        case 'medium': return `2px solid ${c}`;
        case 'thick': return `3px solid ${c}`;
        case 'dashed': return `1px dashed ${c}`;
        case 'dotted': return `1px dotted ${c}`;
        case 'double': return `3px double ${c}`;
        default: return `1px solid ${c}`;
    }
}

export class XlsxEditor extends EditorBase {
    private model: Model = { sheets: [] };
    private activeSheet = 0;
    private sel: Range = { r1: 0, c1: 0, r2: 0, c2: 0 };
    private gridWrap: HTMLElement | null = null;
    private gridCanvas: HTMLElement | null = null;
    private tabsEl: HTMLElement | null = null;
    private nameBoxEl: HTMLInputElement | null = null;
    private formulaEl: HTMLInputElement | null = null;
    private editing: { r: number; c: number } | null = null;
    private statusRef?: (s: string) => void;
    private past: XlsxSheet[][] = [];
    private future: XlsxSheet[][] = [];
    private undoBtn: HTMLButtonElement | null = null;
    private redoBtn: HTMLButtonElement | null = null;
    private readonlyMode: boolean;
    /** 公式栏正在输入（避免外部刷新覆盖用户输入） */
    private formulaTyping = false;

    constructor(hooks: EditorHooks, options: XlsxEditorOptions = {}) {
        super(hooks);
        this.readonlyMode = !!options.readonly;
    }

    async load(bytes: Uint8Array): Promise<void> {
        const parsed = await xlsxToJson(bytes, { assets: 'inline' });
        this.model = {
            metadata: parsed.metadata,
            sheets: (parsed.sheets || []) as XlsxSheet[],
            assets: parsed.assets as unknown[] | undefined
        };
        this.activeSheet = 0;
        // 只补算「文件里没有缓存结果」的公式：文件自带的缓存值是 Excel 算过的，
        // 覆盖它反而会把正确数字改成 0（中文表名跨表求和最典型）。
        this.fillMissingResults();
        this.buildToolbar();
        this.render();
    }

    protected async serialize(): Promise<Uint8Array> {
        // 深拷贝快照：避免 jsonToXlsx(async) 序列化期间模型被再次修改而产出中间态。
        const wb: XlsxWorkbook = { metadata: this.model.metadata, sheets: this.cloneSheets() };
        return jsonToXlsx(wb, this.model.assets ? { assets: this.model.assets as never } : {});
    }

    destroy(): void {
        this.gridWrap = null;
        this.gridCanvas = null;
        this.tabsEl = null;
        this.nameBoxEl = null;
        this.formulaEl = null;
    }

    // ---------- 撤销 / 重做 ----------
    private cloneSheets(): XlsxSheet[] {
        return JSON.parse(JSON.stringify(this.model.sheets)) as XlsxSheet[];
    }
    /** 记录一次「改动前」的快照 */
    private checkpoint(): void {
        this.pushUndo(this.cloneSheets());
    }
    /** 把快照压入撤销栈（无实际改动时不要压） */
    private pushUndo(snapshot: XlsxSheet[]): void {
        this.past.push(snapshot);
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
        if (this.readonlyMode) {
            // 只读模式不提供编辑工具栏
            return;
        }
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

    // ---------- 渲染骨架 ----------
    private render(): void {
        const content = this.contentEl;
        if (!content) return;
        content.innerHTML = '';

        const wrap = h('div', { class: 'oe-xl' });

        // 公式栏
        const nameBox = h('input', { class: 'oe-namebox', readonly: 'readonly', title: '当前单元格' }) as HTMLInputElement;
        const formula = h('input', { class: 'oe-formula', placeholder: '输入内容或以 = 开头输入公式', spellcheck: 'false' }) as HTMLInputElement;
        this.nameBoxEl = nameBox;
        this.formulaEl = formula;
        formula.addEventListener('focus', () => {
            this.formulaTyping = true;
        });
        formula.addEventListener('blur', () => {
            this.formulaTyping = false;
        });
        formula.addEventListener('keydown', (e: KeyboardEvent) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.commitFormulaBar();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this.refreshFormulaBar();
                formula.blur();
            }
        });
        const formulaBar = h('div', { class: 'oe-xl-formula' }, [
            nameBox,
            h('span', { class: 'oe-fx', text: 'fx' }),
            formula
        ]);

        // 网格
        const gridWrap = h('div', { class: 'oe-grid-wrap', tabindex: '0' });
        this.gridWrap = gridWrap;
        gridWrap.addEventListener('keydown', (e: KeyboardEvent) => this.onGridKey(e));
        gridWrap.addEventListener('mousedown', () => {
            if (!this.readonlyMode) gridWrap.focus({ preventScroll: true });
        });
        const canvas = h('div', { class: 'oe-grid-canvas' });
        this.gridCanvas = canvas;
        gridWrap.append(canvas);

        // 底部：工作表标签
        const tabs = h('div', { class: 'oe-sheet-tabs' });
        this.tabsEl = tabs;
        const bottom = h('div', { class: 'oe-xl-bottom' }, [tabs]);

        wrap.append(formulaBar, gridWrap, bottom);
        content.append(wrap);

        this.renderTabs();
        this.paintGrid();
    }

    // ---------- 工作表标签 ----------
    private renderTabs(): void {
        const host = this.tabsEl;
        if (!host) return;
        host.innerHTML = '';
        this.model.sheets.forEach((s, i) => {
            const btn = h('button', {
                class: i === this.activeSheet ? 'active' : '',
                text: s.name,
                title: s.name,
                onclick: () => {
                    if (this.activeSheet === i) return;
                    this.commitEdit();
                    this.activeSheet = i;
                    this.sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
                    this.renderTabs();
                    this.paintGrid();
                }
            });
            host.append(btn);
        });
        if (!this.readonlyMode) {
            host.append(h('button', { class: 'add', title: '新建工作表', text: '+', onclick: () => this.addSheet() }));
        }
    }

    // ---------- 网格尺寸 ----------
    /** 列宽（px）：Excel 列宽单位是「默认字体 0 的宽度」，约 7px/字符 + 5px 内边距 */
    private colWidthPx(sheet: XlsxSheet, c: number): number {
        const w = sheet.columns && sheet.columns[c] ? sheet.columns[c].width : undefined;
        if (!w || !isFinite(w)) return DEFAULT_COL_W;
        return Math.max(16, Math.round(w * 7 + 5));
    }
    /** 行高（px）：磅换像素（96dpi） */
    private rowHeightPx(sheet: XlsxSheet, r: number): number {
        const row = (sheet.rows[r] as XlsxRow | undefined) ?? undefined;
        const hgt = row?.height;
        if (!hgt || !isFinite(hgt)) return DEFAULT_ROW_H;
        return Math.max(14, Math.round(hgt * (96 / 72)));
    }
    /** 图表占用到的最大行列（0 基 + 1） */
    private chartExtent(sheet: XlsxSheet): { cols: number; rows: number } {
        let cols = 0;
        let rows = 0;
        for (const chart of sheet.charts || []) {
            const from = chart.anchor.from;
            const to = chart.anchor.to;
            cols = Math.max(cols, from.col + 1, to ? to.col + 1 : 0);
            rows = Math.max(rows, from.row + 1, to ? to.row + 1 : 0);
        }
        return { cols, rows };
    }

    // ---------- 网格绘制 ----------
    private paintGrid(): void {
        const host = this.gridCanvas;
        const sheet = this.sheet();
        if (!host || !sheet) return;
        host.innerHTML = '';

        const rows = (sheet.rows || []) as XlsxRow[];
        const extent = this.chartExtent(sheet);
        const colCount = Math.max(
            1,
            ...rows.map((r) => (r.cells || []).length),
            (sheet.columns || []).length,
            extent.cols
        );
        const rowCount = Math.max(rows.length, extent.rows);

        // 合并区
        const merges = this.mergesOf(sheet);
        const covered = new Set<string>();
        const anchors = new Map<string, { rs: number; cs: number }>();
        for (const m of merges) {
            anchors.set(m.r1 + ':' + m.c1, { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
            for (let r = m.r1; r <= m.r2; r++) {
                for (let c = m.c1; c <= m.c2; c++) {
                    if (!(r === m.r1 && c === m.c1)) covered.add(r + ':' + c);
                }
            }
        }

        // 冻结窗格
        const fRows = (sheet.frozen && sheet.frozen.rows) || 0;
        const fCols = (sheet.frozen && sheet.frozen.columns) || 0;

        const table = h('table', { class: 'oe-grid' });
        const colgroup = h('colgroup');
        colgroup.append(h('col', { style: `width:${ROW_HEADER_W}px` }));
        for (let c = 0; c < colCount; c++) {
            colgroup.append(h('col', { style: `width:${this.colWidthPx(sheet, c)}px` }));
        }
        table.append(colgroup);

        // 列标行
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

        // 数据行
        const tbody = h('tbody');
        for (let r = 0; r < rowCount; r++) {
            const row = rows[r] as XlsxRow | undefined;
            const tr = h('tr');
            tr.style.height = this.rowHeightPx(sheet, r) + 'px';
            const frozenRow = r < fRows;
            let topPx = HEAD_H;
            for (let k = 0; k < r; k++) topPx += this.rowHeightPx(sheet, k);
            const rn = h('th', { class: 'oe-rownum' + (frozenRow ? ' fz' : ''), text: String(r + 1) });
            if (frozenRow) rn.style.top = topPx + 'px';
            tr.append(rn);

            let leftPx = ROW_HEADER_W;
            for (let c = 0; c < colCount; c++) {
                const cw = this.colWidthPx(sheet, c);
                if (covered.has(r + ':' + c)) {
                    leftPx += cw;
                    continue;
                }
                const td = h('td', { dataset: { r: String(r), c: String(c) } });
                const span = anchors.get(r + ':' + c);
                if (span) {
                    if (span.rs > 1) td.rowSpan = span.rs;
                    if (span.cs > 1) td.colSpan = span.cs;
                }
                if (frozenRow) {
                    td.classList.add('fz-row');
                    td.style.top = topPx + 'px';
                }
                if (c < fCols) {
                    td.classList.add('fz-col');
                    td.style.left = leftPx + 'px';
                }
                leftPx += cw;
                this.paintCell(td, sheet, r, c, row);
                tr.append(td);
            }
            tbody.append(tr);
        }
        table.append(tbody);
        host.append(table);

        // 图表层
        this.paintCharts(host, sheet);

        this.updateSelectionVisual();
        this.updateUndoRedo();
    }

    /** 单元格绘制 */
    private paintCell(td: HTMLElement, sheet: XlsxSheet, r: number, c: number, row: XlsxRow | undefined): void {
        const cell = row?.cells ? (row.cells[c] as XlsxCellObject | null) : null;
        const isEditing = this.editing && this.editing.r === r && this.editing.c === c;
        const el = h('div', { class: 'cell', text: cellDisplay(cell) });
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
            if (st.font?.name) el.style.fontFamily = `"${st.font.name}", "Microsoft YaHei", "PingFang SC", sans-serif`;
            if (st.align) el.style.textAlign = st.align;
            if (st.valign) td.style.verticalAlign = st.valign === 'middle' ? 'middle' : st.valign;
            if (st.fill) el.style.background = st.fill;
            if (st.wrapText) el.classList.add('wrap');
            // 边框
            const b = st.border;
            if (b) {
                if (b.top) td.style.borderTop = borderCss(b.top.style, b.top.color);
                if (b.bottom) td.style.borderBottom = borderCss(b.bottom.style, b.bottom.color);
                if (b.left) td.style.borderLeft = borderCss(b.left.style, b.left.color);
                if (b.right) td.style.borderRight = borderCss(b.right.style, b.right.color);
            }
        }

        // 文本溢出到右侧空单元格（Excel 观感）
        if (!st?.wrapText && (!st?.align || st.align === 'left')) {
            const next = row?.cells ? (row.cells[c + 1] as XlsxCellObject | null) : null;
            if (!next) el.classList.add('spill');
        }

        if (isEditing) {
            el.contentEditable = 'true';
            el.textContent = cell?.formula ? '=' + cell.formula : cellDisplay(cell);
        }
        td.append(el);
        if (!this.readonlyMode) {
            td.addEventListener('mousedown', (e: MouseEvent) => this.onCellMouseDown(e, r, c));
            td.addEventListener('dblclick', () => this.beginEdit(r, c));
        } else {
            td.addEventListener('mousedown', () => {
                this.sel = { r1: r, c1: c, r2: r, c2: c };
                this.updateSelectionVisual();
            });
        }
    }

    // ---------- 图表 ----------
    private paintCharts(host: HTMLElement, sheet: XlsxSheet): void {
        const charts = sheet.charts || [];
        if (!charts.length) return;
        const layer = h('div', { class: 'oe-charts' });
        for (const chart of charts) {
            const rect = this.chartRect(sheet, chart);
            const data = this.chartData(chart);
            // 绘图走共享模块（与 PPTX 同一套内核），这里只把 XLSX 图表转成 ChartSpec
            const spec = xlsxChartToSpec(chart, data);
            const box = h('div', { class: 'oe-chart' });
            box.style.width = Math.round(rect.w) + 'px';
            box.style.height = Math.round(rect.h) + 'px';
            box.innerHTML = renderChartSvg(spec, { width: rect.w, height: rect.h });
            box.style.left = rect.x + 'px';
            box.style.top = rect.y + 'px';
            layer.append(box);
        }
        host.append(layer);
    }

    /** 图表在网格中的位置与尺寸（px） */
    private chartRect(sheet: XlsxSheet, chart: XlsxChart): { x: number; y: number; w: number; h: number } {
        const from = chart.anchor.from;
        const to = chart.anchor.to;
        let x = ROW_HEADER_W;
        for (let c = 0; c < from.col; c++) x += this.colWidthPx(sheet, c);
        x += (from.colOff || 0) / EMU_PER_PX;
        let y = HEAD_H;
        for (let r = 0; r < from.row; r++) y += this.rowHeightPx(sheet, r);
        y += (from.rowOff || 0) / EMU_PER_PX;

        let w = 0;
        let h = 0;
        if (to) {
            let x2 = ROW_HEADER_W;
            for (let c = 0; c < to.col; c++) x2 += this.colWidthPx(sheet, c);
            x2 += (to.colOff || 0) / EMU_PER_PX;
            let y2 = HEAD_H;
            for (let r = 0; r < to.row; r++) y2 += this.rowHeightPx(sheet, r);
            y2 += (to.rowOff || 0) / EMU_PER_PX;
            w = Math.max(60, x2 - x);
            h = Math.max(40, y2 - y);
        } else if (chart.anchor.ext) {
            w = chart.anchor.ext.cx / EMU_PER_PX;
            h = chart.anchor.ext.cy / EMU_PER_PX;
        } else {
            w = 480;
            h = 288;
        }
        return { x, y, w, h };
    }

    /** 图表数据：按引用回查工作簿（改单元格图表跟着变） */
    private chartData(chart: XlsxChart): XlsxChartSeriesData[] {
        const toNumbers = (list: (string | number | null)[] | undefined): (number | null)[] =>
            (list || []).map((v) => (typeof v === 'number' ? v : v === null || v === undefined ? null : Number(v)));
        return (chart.series || []).map((s) => ({
            name: s.name ?? this.refText(s.nameRef) ?? '',
            cats: this.refValues(s.catRef) ?? s.cats ?? [],
            vals: toNumbers(this.refValues(s.valRef) ?? s.vals),
            xVals: s.xRef ? toNumbers(this.refValues(s.xRef) ?? s.xVals) : s.xVals
        }));
    }

    /** 引用 -> 单元格值列表 */
    private refValues(ref?: string): (string | number | null)[] | null {
        if (!ref) return null;
        const parsed = parseRangeRef(ref);
        if (!parsed) return null;
        const sheet = parsed.sheet
            ? this.model.sheets.find((s) => s.name === parsed.sheet)
            : this.sheet();
        if (!sheet) return null;
        const out: (string | number | null)[] = [];
        for (let r = parsed.r1; r <= parsed.r2; r++) {
            const row = (sheet.rows[r] as XlsxRow | undefined) ?? undefined;
            for (let c = parsed.c1; c <= parsed.c2; c++) {
                const cell = row?.cells ? (row.cells[c] as XlsxCellObject | null) : null;
                const info = cellInfo(cell);
                out.push(info.value === null ? null : (info.value as string | number));
            }
        }
        return out;
    }

    /** 引用 -> 首个单元格文本 */
    private refText(ref?: string): string | null {
        const values = this.refValues(ref);
        if (!values || !values.length) return null;
        return values[0] == null ? '' : String(values[0]);
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
            const el = td as HTMLElement;
            const r = Number(el.dataset.r);
            const c = Number(el.dataset.c);
            el.classList.toggle('inrange', this.inSel(r, c));
            el.classList.toggle('active', r === this.sel.r1 && c === this.sel.c1);
        });
        this.refreshFormulaBar();
        this.updateStatus();
    }

    /** 公式栏刷新（选区变化时） */
    private refreshFormulaBar(): void {
        if (this.formulaTyping) return;
        if (this.nameBoxEl) this.nameBoxEl.value = toA1(this.sel.r1, this.sel.c1);
        if (this.formulaEl) {
            const sheet = this.sheet();
            const row = sheet ? (sheet.rows[this.sel.r1] as XlsxRow | undefined) : undefined;
            const cell = row?.cells ? (row.cells[this.sel.c1] as XlsxCellObject | null) : null;
            this.formulaEl.value = cellFormula(cell);
        }
    }

    /** 公式栏回车：写入当前单元格 */
    private commitFormulaBar(): void {
        if (this.readonlyMode) return;
        const input = this.formulaEl;
        if (!input) return;
        const text = input.value;
        this.formulaTyping = false;
        const sheet = this.sheet();
        if (!sheet) return;
        this.writeCell(sheet, this.sel.r1, this.sel.c1, text);
        this.paintGrid();
        this.updateStatus();
    }

    /** 单元格对象 */
    private cellAt(sheet: XlsxSheet, r: number, c: number): XlsxCellObject | null {
        const row = sheet.rows[r] as XlsxRow | undefined;
        return row?.cells ? (row.cells[c] as XlsxCellObject | null) : null;
    }

    private beginEdit(r: number, c: number, initial?: string): void {
        if (this.readonlyMode) return;
        this.editing = { r, c };
        this.paintGrid();
        const td = this.gridWrap?.querySelector(`td[data-r="${r}"][data-c="${c}"]`) as HTMLElement | null;
        const el = td?.querySelector('.cell') as HTMLElement | null;
        if (el) {
            el.focus();
            const range = document.createRange();
            if (initial !== undefined) {
                el.textContent = initial;
                range.selectNodeContents(el);
                range.collapse(false);
            } else {
                range.selectNodeContents(el);
            }
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
            e.preventDefault();
            this.editing = null;
            this.paintGrid();
        }
        e.stopPropagation();
    }

    /** 网格快捷键 */
    private onGridKey(e: KeyboardEvent): void {
        if (this.readonlyMode) return;
        if (this.editing) return;
        const key = e.key;
        // 撤销 / 重做
        if (e.ctrlKey || e.metaKey) {
            const k = key.toLowerCase();
            if (k === 'z') {
                e.preventDefault();
                if (e.shiftKey) this.redo();
                else this.undo();
                return;
            }
            if (k === 'y') {
                e.preventDefault();
                this.redo();
                return;
            }
        }
        if (key === 'ArrowUp') { e.preventDefault(); this.moveActive(-1, 0, e.ctrlKey || e.metaKey); }
        else if (key === 'ArrowDown') { e.preventDefault(); this.moveActive(1, 0, e.ctrlKey || e.metaKey); }
        else if (key === 'ArrowLeft') { e.preventDefault(); this.moveActive(0, -1, e.ctrlKey || e.metaKey); }
        else if (key === 'ArrowRight') { e.preventDefault(); this.moveActive(0, 1, e.ctrlKey || e.metaKey); }
        else if (key === 'Tab') { e.preventDefault(); this.moveActive(0, e.shiftKey ? -1 : 1); }
        else if (key === 'Enter' || key === 'F2') { e.preventDefault(); this.beginEdit(this.sel.r1, this.sel.c1); }
        else if (key === 'Delete' || key === 'Backspace') { e.preventDefault(); this.clearSelection(); }
        else if (key === 'Escape') { e.preventDefault(); this.sel = { r1: this.sel.r1, c1: this.sel.c1, r2: this.sel.r1, c2: this.sel.c1 }; this.updateSelectionVisual(); }
        else if (key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            e.preventDefault();
            this.beginEdit(this.sel.r1, this.sel.c1, key);
        }
    }

    private moveActive(dr: number, dc: number, jump = false): void {
        const sheet = this.sheet();
        if (!sheet) return;
        const rows = (sheet.rows || []) as XlsxRow[];
        const maxR = Math.max(0, rows.length - 1);
        const maxC = Math.max(0, ...rows.map((rw) => Math.max(0, (rw.cells || []).length - 1)));
        let r = this.sel.r1;
        let c = this.sel.c1;
        if (jump) {
            // Ctrl+方向：跳到数据区边缘
            const hasValue = (rr: number, cc: number): boolean => {
                const row = rows[rr] as XlsxRow | undefined;
                const cell = row?.cells ? (row.cells[cc] as XlsxCellObject | null) : null;
                return !!cell && cellDisplay(cell) !== '';
            };
            if (dr !== 0) {
                while (r + dr >= 0 && r + dr <= maxR && hasValue(r + dr, c)) r += dr;
            }
            if (dc !== 0) {
                while (c + dc >= 0 && c + dc <= maxC && hasValue(r, c + dc)) c += dc;
            }
        } else {
            r += dr;
            c += dc;
        }
        r = Math.max(0, Math.min(maxR, r));
        c = Math.max(0, Math.min(maxC, c));
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
        if (sheet && el) {
            this.writeCell(sheet, r, c, el.innerText || '');
        }
        this.paintGrid();
    }

    /** 清空选区内容 */
    private clearSelection(): void {
        const sheet = this.sheet();
        if (!sheet) return;
        const snapshot = this.cloneSheets();
        let changed = false;
        for (let r = this.sel.r1; r <= this.sel.r2; r++) {
            for (let c = this.sel.c1; c <= this.sel.c2; c++) {
                const cell = this.cellAt(sheet, r, c);
                if (!cell) continue;
                if (cell.value === undefined && cell.text === undefined && cell.formula === undefined && !cell.richText) continue;
                delete cell.value;
                delete cell.text;
                delete cell.formula;
                delete cell.result;
                delete cell.richText;
                changed = true;
            }
        }
        if (!changed) return;
        this.pushUndo(snapshot);
        this.paintGrid();
        this.markDirty();
    }

    /** 写入单元格（值或公式） */
    private writeCell(sheet: XlsxSheet, r: number, c: number, raw: string): void {
        const snapshot = this.cloneSheets();
        const input = String(raw).replace(/\u00a0/g, ' ').trim();
        const cell = ensureCell(sheet, r, c);
        const before = cellDisplay(cell);
        // 保留样式，只改内容
        cell.value = undefined;
        cell.text = undefined;
        cell.formula = undefined;
        cell.result = undefined;
        cell.richText = undefined;
        if (input === '') {
            // 清空
        } else if (input.startsWith('=')) {
            cell.formula = input.slice(1);
            try {
                cell.result = evaluateFormula(cell.formula, this.formulaContext(sheet)) as never;
            } catch {
                cell.result = null;
            }
        } else if (/^-?\d+(\.\d+)?$/.test(input)) {
            cell.value = Number(input);
        } else {
            cell.value = input;
        }
        if (before === cellDisplay(cell)) {
            // 内容没变：还原快照，不留空撤销步骤
            this.model.sheets = snapshot;
            return;
        }
        this.pushUndo(snapshot);
        // 沿依赖图增量重算受影响的公式（含跨表）
        this.recomputeDependencies([{ sheet: sheet.name, r, c }]);
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
        this.checkpoint();
        this.eachSelected((cell) => {
            fontOf(cell)[key] = on as never;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setFontSize(size: number): void {
        this.checkpoint();
        this.eachSelected((cell) => {
            fontOf(cell).size = size;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setNumFmt(fmt: string): void {
        this.checkpoint();
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
        if (!sheet) return null;
        return this.cellAt(sheet, this.sel.r1, this.sel.c1);
    }

    private setAlign(a: 'left' | 'center' | 'right'): void {
        this.checkpoint();
        this.eachSelected((cell) => {
            styleOf(cell).align = a;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setFill(color: string): void {
        this.checkpoint();
        this.eachSelected((cell) => {
            styleOf(cell).fill = color;
        });
        this.paintGrid();
        this.markDirty();
    }

    private setFontColor(color: string): void {
        this.checkpoint();
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
        this.shiftRows(sheet, at, 1, 0);
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
        this.shiftRows(sheet, this.sel.r1, -count, 0);
        this.sel = {
            r1: Math.min(this.sel.r1, (sheet.rows as XlsxRow[]).length - 1),
            c1: this.sel.c1,
            r2: Math.min(this.sel.r2, (sheet.rows as XlsxRow[]).length - 1),
            c2: this.sel.c2
        };
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
        this.shiftCols(sheet, at, 1, 0);
        this.paintGrid();
        this.markDirty();
    }
    private deleteCol(): void {
        const sheet = this.sheet();
        if (!sheet) return;
        this.checkpoint();
        const count = this.sel.c2 - this.sel.c1 + 1;
        for (const row of sheet.rows as XlsxRow[]) row.cells?.splice(this.sel.c1, count);
        this.shiftCols(sheet, this.sel.c1, -count, 0);
        this.paintGrid();
        this.markDirty();
    }

    /**
     * 插删行后同步调整合并区与图表锚点
     *
     * 不调整的话，图表会停在原来的行号上（插行后整体偏上），合并区也会错位。
     * Excel 的规则是「受影响区域整体位移，被删掉的行从区域里消失」。
     *
     * @param at - 变动起始行（0 基）
     * @param delta - 位移量：插入为 +n，删除为 -n
     */
    private shiftRows(sheet: XlsxSheet, at: number, delta: number, _unused?: number): void {
        const count = delta < 0 ? -delta : 0;
        const shift = (r: number): number => {
            if (delta > 0) return r >= at ? r + delta : r;
            if (r >= at + count) return r - count;
            if (r >= at) return at - 1; // 落在被删区间：贴到删除点之前
            return r;
        };
        if (sheet.merges?.length) {
            const next: string[] = [];
            for (const m of sheet.merges) {
                const [a, b] = m.split(':');
                const s = fromA1(a);
                const e = fromA1(b || a);
                const r1 = shift(s.r);
                const r2 = shift(e.r);
                if (r2 <= r1) continue; // 塌成一格或被删光：合并区失效
                next.push(`${toA1(r1, s.c)}:${toA1(r2, e.c)}`);
            }
            sheet.merges = next;
        }
        for (const chart of sheet.charts || []) {
            chart.anchor.from.row = shift(chart.anchor.from.row);
            if (chart.anchor.to) chart.anchor.to.row = shift(chart.anchor.to.row);
        }
    }

    /** 插删列后同步调整合并区与图表锚点（规则同 shiftRows） */
    private shiftCols(sheet: XlsxSheet, at: number, delta: number, _unused?: number): void {
        const count = delta < 0 ? -delta : 0;
        const shift = (c: number): number => {
            if (delta > 0) return c >= at ? c + delta : c;
            if (c >= at + count) return c - count;
            if (c >= at) return at - 1;
            return c;
        };
        if (sheet.merges?.length) {
            const next: string[] = [];
            for (const m of sheet.merges) {
                const [a, b] = m.split(':');
                const s = fromA1(a);
                const e = fromA1(b || a);
                const c1 = shift(s.c);
                const c2 = shift(e.c);
                if (c2 <= c1) continue;
                next.push(`${toA1(s.r, c1)}:${toA1(e.r, c2)}`);
            }
            sheet.merges = next;
        }
        for (const chart of sheet.charts || []) {
            chart.anchor.from.col = shift(chart.anchor.from.col);
            if (chart.anchor.to) chart.anchor.to.col = shift(chart.anchor.to.col);
        }
    }

    // ---------- 列宽拖拽 ----------
    private colResize(c: number, e: MouseEvent): void {
        const sheet = this.sheet();
        const host = this.gridWrap;
        if (!sheet || !host) return;
        const cols = host.querySelectorAll('colgroup col');
        const targetCol = cols[c + 1] as HTMLElement | undefined;
        const w0 = (sheet.columns && sheet.columns[c] && sheet.columns[c].width) || 8.43;
        const startX = e.clientX;
        const onMove = (ev: MouseEvent) => {
            const w = Math.max(2, w0 + (ev.clientX - startX) / 7);
            if (targetCol) targetCol.style.width = Math.round(w * 7 + 5) + 'px';
        };
        const onUp = (ev: MouseEvent) => {
            const w = Math.max(2, w0 + (ev.clientX - startX) / 7);
            this.checkpoint();
            if (!sheet.columns) sheet.columns = [];
            if (!sheet.columns[c]) sheet.columns[c] = {};
            sheet.columns[c].width = Math.round(w * 100) / 100;
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            this.paintGrid();
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

    // ---------- 公式求值 ----------
    /** 求值上下文 */
    private formulaContext(sheet: XlsxSheet): FormulaContext {
        return {
            sheetName: sheet.name,
            hasSheet: (name: string) => this.model.sheets.some((s) => s.name === name),
            read: (sheetName: string, row: number, col: number) => {
                const target = this.model.sheets.find((s) => s.name === sheetName);
                const r = target ? (target.rows[row] as XlsxRow | undefined) : undefined;
                const cell = r?.cells ? (r.cells[col] as XlsxCellObject | null) : null;
                const info = cellInfo(cell);
                // 日期在公式里按序列号参与运算（Excel 语义：=B3 得 46033）
                const value =
                    info.value instanceof Date
                        ? Math.round((info.value.getTime() - Date.UTC(1899, 11, 30)) / 86400000)
                        : info.value;
                return { value: value as FormulaValue | null, text: info.text };
            }
        };
    }

    /** 补算「文件里没有缓存结果」的公式（不动已有缓存值） */
    private fillMissingResults(): void {
        for (const sheet of this.model.sheets) {
            const ctx = this.formulaContext(sheet);
            for (const row of sheet.rows as XlsxRow[]) {
                for (const cell of row.cells || []) {
                    if (!cell || cell.formula === undefined) continue;
                    if (cell.result !== undefined && cell.result !== null) continue;
                    try {
                        cell.result = evaluateFormula(cell.formula, ctx) as never;
                    } catch {
                        // 算不出来就保持空，界面显示公式原文
                    }
                }
            }
        }
    }

    /** 依赖图增量重算：从变更单元格出发，重算所有（跨表）依赖它的公式 */
    private recomputeDependencies(changed: { sheet: string; r: number; c: number }[]): void {
        // 引用先解析一次（大表里公式多，逐轮重复解析会明显卡）
        const formulas: {
            sheet: XlsxSheet;
            cell: XlsxCellObject;
            r: number;
            c: number;
            refs: { sheet: string; r1: number; c1: number; r2: number; c2: number }[];
        }[] = [];
        for (const s of this.model.sheets) {
            (s.rows as XlsxRow[]).forEach((row, r) =>
                (row.cells || []).forEach((cell, c) => {
                    if (!cell || cell.formula === undefined) return;
                    const refs = extractFormulaRefs(cell.formula).map((rf) => {
                        const parts = rf.ref.split(':');
                        const a = fromA1(parts[0]);
                        const b = parts[1] ? fromA1(parts[1]) : a;
                        return {
                            sheet: rf.sheet || s.name,
                            r1: Math.min(a.r, b.r),
                            c1: Math.min(a.c, b.c),
                            r2: Math.max(a.r, b.r),
                            c2: Math.max(a.c, b.c)
                        };
                    });
                    formulas.push({ sheet: s, cell, r, c, refs });
                })
            );
        }
        if (!formulas.length) return;

        let frontier = changed.slice();
        const visited = new Set<string>();
        let guard = 0;
        while (frontier.length && guard++ < 100) {
            const next: { sheet: string; r: number; c: number }[] = [];
            for (const ch of frontier) {
                for (const f of formulas) {
                    const key = f.sheet.name + '!' + f.r + ':' + f.c;
                    if (visited.has(key)) continue;
                    const hit = f.refs.some(
                        (rf) =>
                            rf.sheet === ch.sheet &&
                            ch.r >= rf.r1 &&
                            ch.r <= rf.r2 &&
                            ch.c >= rf.c1 &&
                            ch.c <= rf.c2
                    );
                    if (!hit) continue;
                    visited.add(key);
                    const old = f.cell.result;
                    let nextVal: unknown = old;
                    try {
                        nextVal = evaluateFormula(f.cell.formula!, this.formulaContext(f.sheet));
                    } catch {
                        // 算不出来：保留原值，不把正确结果覆盖成 0
                        nextVal = old;
                    }
                    f.cell.result = nextVal as never;
                    if (old !== nextVal) next.push({ sheet: f.sheet.name, r: f.r, c: f.c });
                }
            }
            frontier = next;
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

    // ---------- 状态栏 ----------
    private updateStatus(): void {
        const sheet = this.sheet();
        const cell = this.selActiveCell();
        const addr = `${toA1(this.sel.r1, this.sel.c1)}${this.sel.r1 !== this.sel.r2 || this.sel.c1 !== this.sel.c2 ? ':' + toA1(this.sel.r2, this.sel.c2) : ''}`;
        const stats = this.selectionStats();
        const parts: string[] = [];
        if (this.readonlyMode) parts.push('查看模式（只读）');
        parts.push(`${addr}${sheet ? ' · ' + sheet.name : ''}`);
        if (cell?.formula) parts.push('=' + cell.formula);
        parts.push(`求和=${stats.sum.toLocaleString('zh-CN')} 计数=${stats.count} 平均值=${stats.avg.toLocaleString('zh-CN')}`);
        this.statusRef?.(parts.join('  |  '));
    }

    /** 选区统计：求和 / 计数 / 平均值 */
    private selectionStats(): { sum: number; count: number; avg: number } {
        const sheet = this.sheet();
        let sum = 0;
        let count = 0;
        let nums = 0;
        if (sheet) {
            for (let r = this.sel.r1; r <= this.sel.r2; r++) {
                for (let c = this.sel.c1; c <= this.sel.c2; c++) {
                    const cell = this.cellAt(sheet, r, c);
                    if (!cell) continue;
                    const info = cellInfo(cell);
                    if (info.value === null || info.value === '') continue;
                    count++;
                    if (typeof info.value === 'number') {
                        sum += info.value;
                        nums++;
                    }
                }
            }
        }
        const avg = nums ? Math.round((sum / nums) * 1e6) / 1e6 : 0;
        return { sum: Math.round(sum * 1e6) / 1e6, count, avg };
    }

    setStatusHandler(fn: (s: string) => void): void {
        this.statusRef = fn;
    }
}

export function createXlsxEditor(hooks: EditorHooks, options: XlsxEditorOptions = {}): XlsxEditor {
    return new XlsxEditor(hooks, options);
}
