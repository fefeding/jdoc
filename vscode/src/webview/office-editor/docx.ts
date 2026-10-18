// DOCX 块式编辑器：docxToJson → blocks 模型 ↔ DOM 双向同步 → jsonToDocx 往返。
//
// 每个内容块渲染为独立可编辑区（与 JSON 块一一对应）：
//   heading/paragraph/quote/code → contenteditable（支持字符级富文本）
//   list                        → 多个 contenteditable 列表项
//   table                       → HTML table，单元格 contenteditable，支持增删行列
//   image/pageBreak/hr/textBox  → 只读保留（保证往返不丢）
import { docxToJson, jsonToDocx } from '@fefeding/jdoc';
import type {
    DocxBlock,
    DocxCell,
    DocxCellDetail,
    DocxDocument,
    DocxList,
    DocxMetadata,
    DocxParagraph,
    DocxRun,
    DocxSection,
    DocxTable
} from '@fefeding/jdoc';
import { EditorBase, type EditorHooks } from './framework';
import { h, escapeHtml } from './util';
import { icon } from './icons';

type TextBlock = DocxParagraph | { type: 'heading'; level: number; text?: string; runs?: DocxRun[] } | { type: 'quote'; text?: string; runs?: DocxRun[] } | { type: 'code'; text: string };

interface DocxModel {
    metadata?: DocxMetadata;
    section?: DocxSection;
    blocks: DocxBlock[];
    assets?: unknown[];
}

/** CSS color（rgb()/rgba()/命名）→ #RRGGBB；无法解析返回原值。 */
function normalizeColor(css: string): string {
    const m = css.match(/rgba?\(([^)]+)\)/i);
    if (m) {
        const [r, g, b] = m[1].split(',').map((n) => parseInt(n.trim(), 10));
        if ([r, g, b].every((n) => !isNaN(n))) {
            return '#' + [r, g, b].map((n) => n.toString(16).padStart(2, '0')).join('');
        }
    }
    return css;
}

/** runs/text → contenteditable 的 innerHTML（保留字符级格式）。 */
function runsToInner(runs?: DocxRun[], text?: string): string {
    if (runs && runs.length) {
        return runs
            .map((r) => {
                let s = escapeHtml(r.text || '');
                const styles: string[] = [];
                if (r.color) styles.push(`color:${r.color}`);
                if (r.font) styles.push(`font-family:${r.font}`);
                if (r.size) styles.push(`font-size:${r.size}pt`);
                if (styles.length) s = `<span style="${styles.join(';')}">${s}</span>`;
                if (r.bold) s = `<b>${s}</b>`;
                if (r.italic) s = `<i>${s}</i>`;
                if (r.underline) s = `<u>${s}</u>`;
                if (r.strike) s = `<s>${s}</s>`;
                return s;
            })
            .join('');
    }
    return escapeHtml(text || '');
}

/** 把 contenteditable DOM 解析回 { text, runs? }，保留字符级格式。 */
function readEditable(el: HTMLElement): { text: string; runs?: DocxRun[] } {
    const runs: DocxRun[] = [];
    let hasFormat = false;
    const push = (text: string, fmt: Partial<DocxRun>) => {
        if (!text) return;
        if (fmt.bold || fmt.italic || fmt.underline || fmt.strike || fmt.color || fmt.font || fmt.size) hasFormat = true;
        const last = runs[runs.length - 1];
        if (last && sameFmt(last, fmt)) last.text += text;
        else runs.push({ ...fmt, text });
    };
    const walk = (node: Node, fmt: Partial<DocxRun>) => {
        node.childNodes.forEach((child) => {
            if (child.nodeType === Node.TEXT_NODE) {
                push(child.nodeValue || '', fmt);
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                const e = child as HTMLElement;
                const tag = e.tagName.toLowerCase();
                if (tag === 'br') {
                    push('\n', fmt);
                    return;
                }
                const f: Partial<DocxRun> = { ...fmt };
                if (tag === 'b' || tag === 'strong') f.bold = true;
                else if (tag === 'i' || tag === 'em') f.italic = true;
                else if (tag === 'u') f.underline = true;
                else if (tag === 's' || tag === 'strike' || tag === 'del') f.strike = true;
                const st = e.style;
                if (st) {
                    if (st.color) f.color = normalizeColor(st.color);
                    if (st.fontFamily) f.font = st.fontFamily;
                    if (st.fontSize) f.size = parseInt(st.fontSize, 10);
                    const fw = st.fontWeight;
                    if (fw === 'bold' || (fw && parseInt(fw, 10) >= 600)) f.bold = true;
                    if (st.fontStyle === 'italic') f.italic = true;
                    if (st.textDecoration && st.textDecoration.includes('underline')) f.underline = true;
                    if (st.textDecoration && st.textDecoration.includes('line-through')) f.strike = true;
                }
                walk(e, f);
            }
        });
    };
    walk(el, {});
    const text = (el as HTMLElement).innerText || '';
    return hasFormat && runs.length ? { text, runs } : { text };
}

function sameFmt(a: DocxRun, b: Partial<DocxRun>): boolean {
    return (
        !!a.bold === !!b.bold &&
        !!a.italic === !!b.italic &&
        !!a.underline === !!b.underline &&
        !!a.strike === !!b.strike &&
        (a.color || '') === (b.color || '') &&
        (a.font || '') === (b.font || '') &&
        (a.size || 0) === (b.size || 0)
    );
}

/** 取单元格的可编辑文本与 runs（兼容 string / runs[] / cellDetail）。 */
function cellText(cell: DocxCell | null | undefined): { text: string; runs?: DocxRun[] } {
    if (cell == null) return { text: '' };
    if (typeof cell === 'string') return { text: cell };
    if (Array.isArray(cell)) return { text: cell.map((r) => r.text || '').join(''), runs: cell };
    // DocxCellDetail
    if (cell.runs && cell.runs.length) return { text: cell.runs.map((r) => r.text || '').join(''), runs: cell.runs };
    if (cell.blocks && cell.blocks.length) {
        const p = cell.blocks.find((b) => b.type === 'paragraph') as DocxParagraph | undefined;
        if (p) return { text: p.text || (p.runs || []).map((r) => r.text || '').join('') || '', runs: p.runs };
    }
    return { text: cell.text || '' };
}

function isCellDetail(cell: DocxCell | null | undefined): cell is DocxCellDetail {
    return !!cell && typeof cell === 'object' && !Array.isArray(cell) && (cell as DocxCellDetail).type === 'cell';
}

function alignStyle(align?: string): string {
    if (!align || align === 'left') return '';
    return `text-align:${align};`;
}

export class DocxEditor extends EditorBase {
    private model: DocxModel = { blocks: [] };
    private sheetEl: HTMLElement | null = null;
    private current = -1; // 选中块索引

    async load(bytes: Uint8Array): Promise<void> {
        const parsed = await docxToJson(bytes, { assets: 'inline' });
        this.model = {
            metadata: parsed.metadata,
            section: parsed.section,
            blocks: parsed.blocks || [],
            assets: parsed.assets as unknown[] | undefined
        };
        this.buildToolbar();
        this.renderAll();
    }

    protected async serialize(): Promise<Uint8Array> {
        const doc: DocxDocument = {
            metadata: this.model.metadata,
            section: this.model.section,
            blocks: this.model.blocks
        };
        return jsonToDocx(doc, this.model.assets ? { assets: this.model.assets as never } : {});
    }

    destroy(): void {
        this.sheetEl = null;
    }

    // ---------- 工具栏 ----------
    private buildToolbar(): void {
        const bar = this.toolbarEl;
        if (!bar) return;
        bar.innerHTML = '';
        const sel = h('select', { class: 'oe-select', title: '块类型', onchange: (e: Event) => this.setBlockType((e.target as HTMLSelectElement).value) });
        const opts: [string, string][] = [
            ['paragraph', '正文'], ['h1', '标题 1'], ['h2', '标题 2'], ['h3', '标题 3'],
            ['quote', '引用'], ['code', '代码']
        ];
        for (const [v, label] of opts) sel.append(h('option', { value: v, text: label }));
        bar.append(sel, sep());

        const fmtBtn = (cmd: string, ic: string, title: string) =>
            h('button', { class: 'oe-tb', title, html: icon(ic), onclick: () => this.exec(cmd) });
        bar.append(
            fmtBtn('bold', 'bold', '加粗 (Ctrl+B)'),
            fmtBtn('italic', 'italic', '斜体 (Ctrl+I)'),
            fmtBtn('underline', 'underline', '下划线 (Ctrl+U)'),
            fmtBtn('strikeThrough', 'strike', '删除线'),
            sep()
        );
        const alignBtn = (a: string, ic: string, title: string) =>
            h('button', { class: 'oe-tb', title, html: icon(ic), onclick: () => this.setAlign(a) });
        bar.append(
            alignBtn('left', 'alignLeft', '左对齐'),
            alignBtn('center', 'alignCenter', '居中'),
            alignBtn('right', 'alignRight', '右对齐'),
            alignBtn('justify', 'alignJustify', '两端对齐'),
            sep()
        );
        bar.append(
            h('button', { class: 'oe-tb', title: '无序列表', html: icon('listUl'), onclick: () => this.insertBlock({ type: 'list', ordered: false, items: [''] }) }),
            h('button', { class: 'oe-tb', title: '有序列表', html: icon('listOl'), onclick: () => this.insertBlock({ type: 'list', ordered: true, items: [''] }) }),
            h('button', { class: 'oe-tb', title: '插入表格', html: icon('table'), onclick: () => this.insertTable() }),
            h('button', { class: 'oe-tb', title: '分隔线', html: icon('hr'), onclick: () => this.insertBlock({ type: 'hr' }) }),
            h('button', { class: 'oe-tb', title: '分页符', html: icon('pageBreak'), onclick: () => this.insertBlock({ type: 'pageBreak' }) }),
            sep()
        );
        bar.append(
            h('button', { class: 'oe-tb', title: '上移块', html: icon('up'), onclick: () => this.moveBlock(-1) }),
            h('button', { class: 'oe-tb', title: '下移块', html: icon('down'), onclick: () => this.moveBlock(1) }),
            h('button', { class: 'oe-tb', title: '删除块', html: icon('trash'), onclick: () => this.deleteBlock() }),
            h('span', { class: 'oe-spacer' })
        );
    }

    // ---------- 渲染 ----------
    private renderAll(): void {
        const content = this.contentEl;
        if (!content) return;
        content.innerHTML = '';
        const doc = h('div', { class: 'oe-doc' });
        const sheet = h('div', { class: 'sheet' });
        this.sheetEl = sheet;
        this.model.blocks.forEach((block, i) => {
            sheet.append(this.renderBlock(block, i));
            sheet.append(this.renderInsertBar(i));
        });
        if (!this.model.blocks.length) {
            sheet.append(h('div', { class: 'oe-hint', text: '空文档。用工具栏插入内容，或点击下方 + 号。' }));
            sheet.append(this.renderInsertBar(-1));
        }
        doc.append(sheet);
        content.append(doc);
    }

    private renderInsertBar(afterIdx: number): HTMLElement {
        const bar = h('div', { class: 'oe-insert' });
        const add = (label: string, fn: () => void) => h('button', { type: 'button', text: label, onclick: fn });
        bar.append(
            add('+ 段落', () => this.insertBlock({ type: 'paragraph', text: '' }, afterIdx)),
            add('+ 标题', () => this.insertBlock({ type: 'heading', level: 2, text: '' }, afterIdx)),
            add('+ 表格', () => this.insertTable(afterIdx)),
            add('+ 列表', () => this.insertBlock({ type: 'list', ordered: false, items: [''] }, afterIdx))
        );
        return bar;
    }

    private renderBlock(block: DocxBlock, index: number): HTMLElement {
        const dataset: Record<string, string> = { type: block.type, idx: String(index) };
        if (block.type === 'heading') dataset.level = String((block as { level: number }).level);
        const wrap = h('div', { class: 'oe-block' + (index === this.current ? ' sel' : ''), dataset });
        wrap.addEventListener('mousedown', () => this.select(index));
        const grip = h('div', { class: 'oe-grip' });
        grip.append(
            h('button', { title: '上移', html: icon('up'), onclick: (e: Event) => { e.stopPropagation(); this.moveBlock(-1); } }),
            h('button', { title: '下移', html: icon('down'), onclick: (e: Event) => { e.stopPropagation(); this.moveBlock(1); } }),
            h('button', { title: '删除', html: icon('trash'), onclick: (e: Event) => { e.stopPropagation(); this.deleteBlock(); } })
        );
        wrap.append(grip);
        const body = this.renderBlockBody(block, index);
        wrap.append(body);
        return wrap;
    }

    private renderBlockBody(block: DocxBlock, index: number): HTMLElement {
        switch (block.type) {
            case 'heading':
            case 'paragraph':
            case 'quote':
                return this.makeEditable(block as TextBlock, index, block.type as 'heading' | 'paragraph' | 'quote');
            case 'code':
                return this.makeEditable(block as unknown as TextBlock, index, 'code');
            case 'list':
                return this.renderList(block as DocxList, index);
            case 'table':
                return this.renderTable(block as DocxTable, index);
            case 'image': {
                const img = block as { src: unknown; caption?: string; width?: number; height?: number };
                const src = typeof img.src === 'string' ? img.src : '';
                const el = h('img', { src, alt: img.caption || 'image' });
                if (img.width) el.style.maxWidth = img.width + 'pt';
                const box = h('div', {}, [el]);
                if (img.caption) box.append(h('div', { class: 'cap', text: img.caption }));
                return box;
            }
            case 'hr':
                return h('div', {}, [h('hr')]);
            case 'pageBreak':
                return h('div', { text: '— 分页符 —' });
            case 'textBox': {
                const tb = block as { paragraphs: DocxBlock[]; fill?: string; stroke?: string };
                const box = h('div', { style: `border:1px solid ${tb.stroke || '#cbd5e1'};background:${tb.fill || '#f8fafc'};border-radius:6px;padding:8px 12px;` });
                for (const p of tb.paragraphs || []) {
                    const text = (p as DocxParagraph).text || ((p as DocxParagraph).runs || []).map((r) => r.text || '').join('');
                    box.append(h('div', { text: text || '' }));
                }
                return box;
            }
            default:
                return h('div', { text: (block as { text?: string }).text || '' });
        }
    }

    private renderList(list: DocxList, index: number): HTMLElement {
        const box = h('div', { dataset: { ordered: list.ordered ? '1' : '0' } });
        (list.items || []).forEach((item, ii) => {
            const row = h('div', { class: 'oe-li' });
            const mk = h('span', { class: 'mk', text: list.ordered ? `${ii + 1}.` : '•' });
            const { text, runs } = typeof item === 'string' ? { text: item, runs: undefined } : { text: (item as DocxRun[]).map((r) => r.text || '').join(''), runs: item as DocxRun[] };
            const ed = h('div', { class: 'oe-edit', contenteditable: 'true', html: runsToInner(runs, text) });
            this.bindEditable(ed, () => {
                const r = readEditable(ed);
                if (r.runs) (list.items[ii] as unknown as DocxRun[]) = r.runs as unknown as DocxRun[];
                else list.items[ii] = r.text;
            });
            row.append(mk, ed);
            box.append(row);
        });
        // 列表项操作
        const ops = h('div', { style: 'display:flex;gap:6px;margin-top:4px;' });
        ops.append(
            h('button', { class: 'oe-tb', title: '添加列表项', html: icon('plus'), onclick: () => { (list.items as string[]).push(''); this.renderAll(); this.markDirty(); } }),
            h('button', { class: 'oe-tb', title: list.ordered ? '改为无序' : '改为有序', html: icon(list.ordered ? 'listUl' : 'listOl'), onclick: () => { list.ordered = !list.ordered; this.renderAll(); this.markDirty(); } })
        );
        box.append(ops);
        // 点击列表也选中该块
        box.addEventListener('mousedown', () => this.select(index));
        return box;
    }

    private renderTable(table: DocxTable, index: number): HTMLElement {
        const box = h('div', { dataset: { table: '1' } });
        const tbl = h('table', { class: 'oe-tbl' });
        const rows = table.rows || [];
        rows.forEach((row, ri) => {
            const tr = h('tr');
            (row || []).forEach((cell, ci) => {
                const isHead = !!table.header && ri === 0;
                const cellEl = isHead ? h('th') : h('td');
                const detail = isCellDetail(cell) ? cell : null;
                if (detail?.colspan) cellEl.colSpan = detail.colspan;
                if (detail?.rowspan) cellEl.rowSpan = detail.rowspan;
                if (detail?.fill) cellEl.style.background = detail.fill;
                const { text, runs } = cellText(cell);
                const ed = h('div', { class: 'oe-edit', contenteditable: 'true', html: runsToInner(runs, text) });
                this.bindEditable(ed, () => this.writeCell(table, ri, ci, cell, ed));
                cellEl.append(ed);
                tr.append(cellEl);
            });
            tbl.append(tr);
        });
        box.append(tbl);
        // 表格结构工具
        const ops = h('div', { style: 'display:flex;gap:6px;margin:4px 0;' });
        ops.append(
            h('button', { class: 'oe-tb', title: '在下方添加行', html: icon('plus'), onclick: () => this.tableAddRow(table, index) }),
            h('button', { class: 'oe-tb', title: '删除最后一行', html: icon('trash'), onclick: () => this.tableDelRow(table, index) }),
            h('button', { class: 'oe-tb', title: '添加列', html: icon('down'), onclick: () => this.tableAddCol(table, index) }),
            h('button', { class: 'oe-tb', title: '删除最后一列', html: icon('up'), onclick: () => this.tableDelCol(table, index) })
        );
        box.append(ops);
        box.addEventListener('mousedown', () => this.select(index));
        return box;
    }

    private writeCell(table: DocxTable, ri: number, ci: number, original: DocxCell | null | undefined, ed: HTMLElement): void {
        const r = readEditable(ed);
        const rows = table.rows as DocxCell[][];
        if (!rows[ri]) return;
        if (typeof original === 'string' || original == null) {
            rows[ri][ci] = r.text;
        } else if (Array.isArray(original)) {
            rows[ri][ci] = (r.runs || [{ text: r.text }]) as DocxRun[];
        } else if (isCellDetail(original)) {
            original.text = r.text;
            if (r.runs) original.runs = r.runs;
            delete original.blocks;
            rows[ri][ci] = original;
        }
        this.markDirty();
    }

    private makeEditable(block: TextBlock, index: number, kind: 'paragraph' | 'heading' | 'quote' | 'code'): HTMLElement {
        const align = (block as DocxParagraph).align;
        const ed = h('div', {
            class: 'oe-edit',
            contenteditable: 'true',
            style: alignStyle(align),
            html: runsToInner((block as { runs?: DocxRun[] }).runs, (block as { text?: string }).text),
            dataset: { ph: kind === 'code' ? '代码…' : kind === 'heading' ? '标题' : kind === 'quote' ? '引用' : '输入正文…' }
        });
        this.bindEditable(ed, () => {
            const r = readEditable(ed);
            if (r.runs) {
                (block as { runs?: DocxRun[] }).runs = r.runs;
                delete (block as { text?: string }).text;
            } else {
                (block as { text?: string }).text = r.text;
                delete (block as { runs?: DocxRun[] }).runs;
            }
            this.syncEmptyState(ed);
        });
        ed.addEventListener('focus', () => this.select(index));
        this.syncEmptyState(ed);
        return ed;
    }

    private bindEditable(ed: HTMLElement, commit: () => void): void {
        ed.addEventListener('blur', commit);
        ed.addEventListener('input', () => {
            this.syncEmptyState(ed);
            this.markDirty();
        });
    }

    private syncEmptyState(ed: HTMLElement): void {
        const empty = (ed.innerText || '').trim() === '' && !ed.querySelector('img');
        if (empty) ed.dataset.empty = '1';
        else delete ed.dataset.empty;
    }

    // ---------- 选中 / 命令 ----------
    private select(index: number): void {
        if (this.current === index) return;
        this.current = index;
        this.sheetEl?.querySelectorAll('.oe-block').forEach((el) => {
            el.classList.toggle('sel', Number((el as HTMLElement).dataset.idx) === index);
        });
    }

    private exec(cmd: string): void {
        document.execCommand(cmd);
        const ed = document.activeElement as HTMLElement | null;
        if (ed && ed.classList.contains('oe-edit')) {
            this.syncEmptyState(ed);
            this.markDirty();
        }
    }

    private currentBlock(): DocxBlock | null {
        return this.current >= 0 ? this.model.blocks[this.current] : null;
    }

    private setBlockType(value: string): void {
        const b = this.currentBlock();
        if (!b) return;
        if (value === 'quote' || value === 'code' || value === 'paragraph') {
            const text = blockText(b);
            const nb: DocxBlock = value === 'code' ? { type: 'code', text } : value === 'quote' ? { type: 'quote', text } : { type: 'paragraph', text, align: (b as DocxParagraph).align };
            this.model.blocks[this.current] = nb;
        } else if (value.startsWith('h')) {
            const level = parseInt(value.slice(1), 10) as 1 | 2 | 3 | 4 | 5 | 6;
            const text = blockText(b);
            this.model.blocks[this.current] = { type: 'heading', level, text };
        }
        this.renderAll();
        this.markDirty();
    }

    private setAlign(a: string): void {
        const b = this.currentBlock();
        if (!b || (b.type !== 'paragraph' && b.type !== 'heading' && b.type !== 'quote')) return;
        (b as DocxParagraph).align = a as DocxParagraph['align'];
        // 局部更新样式，避免重建丢失焦点
        const el = this.sheetEl?.querySelector(`.oe-block[data-idx="${this.current}"] .oe-edit`) as HTMLElement | null;
        if (el) el.style.textAlign = a === 'left' ? '' : a;
        this.markDirty();
    }

    private insertBlock(block: DocxBlock, after = this.current): void {
        const at = after < 0 ? this.model.blocks.length : after + 1;
        this.model.blocks.splice(at, 0, block);
        this.current = at;
        this.renderAll();
        this.markDirty();
        this.focusBlock(at);
    }

    private insertTable(after = this.current): void {
        const rows: DocxCell[][] = [
            ['列 A', '列 B', '列 C'],
            ['', '', ''],
            ['', '', '']
        ];
        this.insertBlock({ type: 'table', header: true, rows }, after);
    }

    private moveBlock(dir: number): void {
        const i = this.current;
        const j = i + dir;
        if (i < 0 || j < 0 || j >= this.model.blocks.length) return;
        const [b] = this.model.blocks.splice(i, 1);
        this.model.blocks.splice(j, 0, b);
        this.current = j;
        this.renderAll();
        this.markDirty();
    }

    private deleteBlock(): void {
        if (this.current < 0) return;
        this.model.blocks.splice(this.current, 1);
        this.current = -1;
        this.renderAll();
        this.markDirty();
    }

    private focusBlock(index: number): void {
        window.setTimeout(() => {
            const ed = this.sheetEl?.querySelector(`.oe-block[data-idx="${index}"] .oe-edit`) as HTMLElement | null;
            ed?.focus();
        }, 0);
    }

    // ---------- 表格结构 ----------
    private tableAddRow(table: DocxTable, index: number): void {
        const cols = table.rows[0]?.length || 1;
        (table.rows as DocxCell[][]).push(new Array(cols).fill(''));
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
    private tableDelRow(table: DocxTable, index: number): void {
        if (table.rows.length <= 1) return;
        (table.rows as DocxCell[][]).pop();
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
    private tableAddCol(table: DocxTable, index: number): void {
        for (const row of table.rows) (row as DocxCell[]).push('');
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
    private tableDelCol(table: DocxTable, index: number): void {
        const cols = table.rows[0]?.length || 0;
        if (cols <= 1) return;
        for (const row of table.rows) (row as DocxCell[]).pop();
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
}

function blockText(b: DocxBlock): string {
    const anyB = b as { text?: string; runs?: DocxRun[]; items?: unknown[] };
    if (anyB.text != null) return anyB.text;
    if (anyB.runs) return anyB.runs.map((r) => r.text || '').join('');
    if (anyB.items) return (anyB.items as (string | DocxRun[])[]).map((it) => (typeof it === 'string' ? it : it.map((r) => r.text || '').join(''))).join('\n');
    return '';
}

function sep(): HTMLElement {
    return h('span', { class: 'oe-sep' });
}

export function createDocxEditor(hooks: EditorHooks): DocxEditor {
    return new DocxEditor(hooks);
}
