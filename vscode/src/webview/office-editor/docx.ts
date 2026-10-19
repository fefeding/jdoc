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
import { paginate, waitImages, A4_PORTRAIT, type PageMetrics } from './paging';
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
                // 只读装饰（如标题章节号）不产出 run，否则会被当成正文存回模型
                if (e.getAttribute('contenteditable') === 'false') return;
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
                else if (tag === 'font') {
                    const sz = e.getAttribute('size');
                    if (sz && FONT_SIZE_MAP[sz]) f.size = FONT_SIZE_MAP[sz];
                    const col = e.getAttribute('color');
                    if (col) f.color = normalizeColor(col);
                    const face = e.getAttribute('face');
                    if (face) f.font = face;
                }
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
    const text = plainTextOf(el);
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

const FONTS = ['默认', '宋体', '微软雅黑', '黑体', 'Arial', 'Helvetica', 'Times New Roman', 'Georgia', 'Courier New', 'PingFang SC'];
// execCommand fontSize 的 1-7 档 → 磅值（用于 <font size> 与磅值互转）
const FONT_SIZE_MAP: Record<string, number> = { '1': 8, '2': 10, '3': 12, '4': 14, '5': 18, '6': 24, '7': 36 };
function ptToN(pt: number): string {
    for (const [n, p] of Object.entries(FONT_SIZE_MAP)) if (p === pt) return n;
    return '3';
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

/** DocxSection -> 纸张设置；缺省 A4 纵向 1 英寸页边距 */
function metricsOf(section: DocxSection | undefined): PageMetrics {
    if (!section || !(section.width > 0) || !(section.height > 0)) return { ...A4_PORTRAIT };
    return {
        width: section.width,
        height: section.height,
        marginTop: section.marginTop,
        marginBottom: section.marginBottom,
        marginLeft: section.marginLeft,
        marginRight: section.marginRight
    };
}

/** 标题章节号 span 的类名（解析自 numbering，只展示不进模型） */
const H_NUM_CLASS = 'oe-h-num';

/** 去掉只读装饰（章节号）后的纯文本：contenteditable 的 innerText 含装饰节点 */
function plainTextOf(el: HTMLElement): string {
    const clone = el.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('[contenteditable="false"]').forEach((node) => node.remove());
    return clone.innerText || '';
}

export class DocxEditor extends EditorBase {
    private model: DocxModel = { blocks: [] };
    private docEl: HTMLElement | null = null;
    private current = -1; // 选中块索引
    private past: DocxBlock[][] = [];
    private future: DocxBlock[][] = [];
    private undoBtn: HTMLButtonElement | null = null;
    private redoBtn: HTMLButtonElement | null = null;
    private fmtBtns: Record<string, HTMLButtonElement> = {};
    private savedRange: Range | null = null;
    private selHandler: (() => void) | null = null;
    /** 重渲染代数：丢弃过期的图片重排结果，防异步竞态 */
    private renderGen = 0;

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
        this.selHandler = () => {
            this.saveSel();
            this.updateFmtState();
        };
        document.addEventListener('selectionchange', this.selHandler);
    }

    protected async serialize(): Promise<Uint8Array> {
        // 深拷贝快照：jsonToDocx 为 async，序列化期间模型可能被再次修改，
        // 直接传 this.model.blocks 引用会产出中间态字节。
        const doc: DocxDocument = {
            metadata: this.model.metadata,
            section: this.model.section,
            blocks: this.cloneBlocks()
        };
        return jsonToDocx(doc, this.model.assets ? { assets: this.model.assets as never } : {});
    }

    destroy(): void {
        this.docEl = null;
        if (this.selHandler) document.removeEventListener('selectionchange', this.selHandler);
        this.selHandler = null;
    }

    // ---------- 工具栏 ----------
    private buildToolbar(): void {
        const bar = this.toolbarEl;
        if (!bar) return;
        bar.innerHTML = '';
        this.undoBtn = h('button', { class: 'oe-tb', title: '撤销 (Ctrl+Z)', html: icon('undo'), onmousedown: (e: Event) => e.preventDefault(), onclick: () => this.undo() });
        this.redoBtn = h('button', { class: 'oe-tb', title: '重做 (Ctrl+Shift+Z)', html: icon('redo'), onmousedown: (e: Event) => e.preventDefault(), onclick: () => this.redo() });
        bar.append(this.undoBtn, this.redoBtn, sep());
        const sel = h('select', { class: 'oe-select', title: '块类型', onchange: (e: Event) => this.setBlockType((e.target as HTMLSelectElement).value) });
        const opts: [string, string][] = [
            ['paragraph', '正文'], ['h1', '标题 1'], ['h2', '标题 2'], ['h3', '标题 3'],
            ['quote', '引用'], ['code', '代码']
        ];
        for (const [v, label] of opts) sel.append(h('option', { value: v, text: label }));
        bar.append(sel, sep());

        const fmtBtn = (cmd: string, ic: string, title: string, key: string) => {
            const b = h('button', { class: 'oe-tb', title, html: icon(ic), onmousedown: (e: Event) => e.preventDefault(), onclick: () => this.exec(cmd) });
            this.fmtBtns[key] = b as HTMLButtonElement;
            return b;
        };
        const fontSel = h('select', { class: 'oe-select', title: '字体', style: 'width:104px', onchange: (e: Event) => this.execWith('fontName', (e.target as HTMLSelectElement).value) });
        for (const f of FONTS) fontSel.append(h('option', { value: f, text: f }));
        const sizeSel = h('select', { class: 'oe-select', title: '字号', style: 'width:62px', onchange: (e: Event) => this.execFontSize(parseInt((e.target as HTMLSelectElement).value, 10)) });
        for (const p of [8, 9, 10, 12, 14, 16, 18, 24, 36]) sizeSel.append(h('option', { value: String(p), text: String(p) }));
        bar.append(
            fmtBtn('bold', 'bold', '加粗 (Ctrl+B)', 'bold'),
            fmtBtn('italic', 'italic', '斜体 (Ctrl+I)', 'italic'),
            fmtBtn('underline', 'underline', '下划线 (Ctrl+U)', 'underline'),
            fmtBtn('strikeThrough', 'strike', '删除线', 'strike'),
            h('span', { class: 'oe-sep' }),
            fontSel,
            sizeSel,
            h('input', { type: 'color', class: 'oe-color', title: '文字颜色', value: '#111827', oninput: (e: Event) => this.execWith('foreColor', (e.target as HTMLInputElement).value) }),
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
            h('button', { class: 'oe-tb', title: '插入图片', html: icon('image'), onclick: () => this.pickImage() }),
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
        this.docEl = doc;
        // 先把全部块渲染成游离节点，再按纸张高度装箱成多页（与查看模式观感一致）
        const nodes: HTMLElement[] = [];
        this.model.blocks.forEach((block, i) => {
            nodes.push(this.renderBlock(block, i));
            nodes.push(this.renderInsertBar(i));
        });
        if (!this.model.blocks.length) {
            nodes.push(h('div', { class: 'oe-hint', text: '空文档。用工具栏插入内容，或点击下方 + 号。' }));
            nodes.push(this.renderInsertBar(-1));
        }
        doc.appendChild(this.buildPages(nodes));
        content.append(doc);
        this.updateUndoRedo();
        // 分页已就绪，先刷新大纲；图片解码完成后再重排一次（重排内部会再通知）
        this.notifyContentChanged();
        void this.repaginate();
    }

    /** 按纸张设置把游离节点装箱成多页 */
    private buildPages(nodes: HTMLElement[]): HTMLElement {
        return paginate(nodes, {
            metrics: metricsOf(this.model.section),
            pageClass: 'oe-page oe-page-edit',
            bodyClass: 'oe-page-body',
            pageNumber: true,
            mount: this.contentEl
        });
    }

    /**
     * 图片解码完成后按最终高度重排分页
     *
     * 首轮分页时图片高度为 0，会把图表当成零高度块塞进当前页。必须在图片就绪后
     * 再切一次；重排只搬动节点不改内容，所以不会打断用户正在进行的编辑。
     */
    private async repaginate(): Promise<void> {
        const doc = this.docEl;
        if (!doc || !doc.isConnected) return;
        const my = ++this.renderGen;
        const pages = doc.firstElementChild as HTMLElement | null;
        if (!pages) return;
        await waitImages(pages);
        if (!doc.isConnected || my !== this.renderGen) return;
        const nodes = Array.from(pages.querySelectorAll<HTMLElement>(':scope > .oe-page > .oe-page-body > *'));
        if (!nodes.length) return;
        doc.replaceChild(this.buildPages(nodes), pages);
        this.notifyContentChanged();
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
        // 显式分页符：交给分页器在此处断页（标记元素自身不渲染）
        if (block.type === 'pageBreak') dataset.oePagebreak = '1';
        const wrap = h('div', { class: 'oe-block' + (index === this.current ? ' sel' : ''), dataset });
        wrap.addEventListener('mousedown', () => this.select(index));
        // 分页符块本身不占版面，只留一个不可见的锚点给分页器识别
        if (block.type === 'pageBreak') return wrap;
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
            case 'toc':
                return this.renderToc(block as { entries?: { level?: number; text?: string; page?: number }[] });
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

    /**
     * 目录块 -> 只读渲染
     *
     * 目录条目是 TOC 域的缓存结果（含 Word 自己算的页码），没有可编辑的正文形态。
     * 这里按层级缩进 + 点线导引 + 右对齐页码还原观感，但不进 contenteditable，
     * 避免用户改了条目却仍以旧值往返写回。
     */
    private renderToc(block: { entries?: { level?: number; text?: string; page?: number }[] }): HTMLElement {
        const box = h('div', { class: 'oe-toc' });
        for (const entry of block.entries || []) {
            const level = Math.min(Math.max(entry.level || 1, 1), 3);
            const row = h('div', { class: 'oe-toc-row', dataset: { level: String(level) } });
            row.append(h('span', { class: 'oe-toc-text', text: entry.text || '' }));
            row.append(h('span', { class: 'oe-toc-dots' }));
            if (entry.page !== undefined) row.append(h('span', { class: 'oe-toc-page', text: String(entry.page) }));
            box.append(row);
        }
        return box;
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
        // jdoc 的表格是「压缩表示」：被 colspan/rowspan 覆盖的位置不产 cell，数组下标≠列号。
        // 因此渲染必须按列号追踪：跳过被上方 rowspan 占用的列，再按 colspan 占列。
        const occupied = new Set<string>(); // 'row:col' 被上方 rowspan 占用
        rows.forEach((row, ri) => {
            const tr = h('tr');
            let col = 0;
            (row || []).forEach((cell, ci) => {
                while (occupied.has(ri + ':' + col)) col++;
                const isHead = !!table.header && ri === 0;
                const cellEl = isHead ? h('th') : h('td');
                const detail = isCellDetail(cell) ? cell : null;
                const cs = detail?.colspan || 1;
                const rs = detail?.rowspan || 1;
                if (cs > 1) cellEl.colSpan = cs;
                if (rs > 1) {
                    cellEl.rowSpan = rs;
                    for (let k = 1; k < rs; k++) for (let c2 = 0; c2 < cs; c2++) occupied.add(ri + k + ':' + (col + c2));
                }
                if (detail?.fill) cellEl.style.background = detail.fill;
                cellEl.dataset.r = String(ri);
                cellEl.dataset.c = String(ci);
                const { text, runs } = cellText(cell);
                const ed = h('div', { class: 'oe-edit', contenteditable: 'true', html: runsToInner(runs, text) });
                this.bindEditable(ed, () => this.writeCell(table, ri, ci, cell, ed));
                cellEl.append(ed);
                tr.append(cellEl);
                col += cs;
            });
            tbl.append(tr);
        });
        box.append(tbl);
        // 表格结构工具（增删行列）
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
            dataset: { ph: kind === 'code' ? '代码…' : kind === 'heading' ? '标题' : kind === 'quote' ? '引用' : '输入正文…' }
        });
        // 章节号是解析自 numbering 的只读装饰（contenteditable=false），
        // 既能显示 Word/WPS 的编号，又不会被编辑进正文或存回模型
        const number = kind === 'heading' ? (block as { number?: string }).number : undefined;
        if (number) {
            ed.appendChild(h('span', { class: H_NUM_CLASS, contenteditable: 'false', text: number }));
            ed.appendChild(document.createTextNode(' '));
        }
        ed.insertAdjacentHTML('beforeend', runsToInner((block as { runs?: DocxRun[] }).runs, (block as { text?: string }).text));
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
        (ed as { __commit?: () => void }).__commit = commit;
        ed.addEventListener('blur', () => {
            delete ed.dataset.ck;
            commit();
        });
        ed.addEventListener('input', () => {
            this.syncEmptyState(ed);
            if (!ed.dataset.ck) {
                this.checkpoint();
                ed.dataset.ck = '1';
            }
            this.markDirty();
        });
    }

    private syncEmptyState(ed: HTMLElement): void {
        const empty = (ed.innerText || '').trim() === '' && !ed.querySelector('img');
        if (empty) ed.dataset.empty = '1';
        else delete ed.dataset.empty;
    }

    // ---------- 撤销 / 重做 ----------
    private cloneBlocks(): DocxBlock[][] {
        return JSON.parse(JSON.stringify(this.model.blocks)) as DocxBlock[][];
    }
    private checkpoint(): void {
        this.past.push(this.cloneBlocks());
        if (this.past.length > 200) this.past.shift();
        this.future = [];
        this.updateUndoRedo();
    }
    undo(): void {
        if (!this.past.length) return;
        this.future.push(this.cloneBlocks());
        this.model.blocks = this.past.pop()!;
        this.renderAll();
        this.updateUndoRedo();
        // 历史栈见底 = 回到导入态：通知 host 放弃编辑（bytes=null）而非回传相同字节
        if (this.past.length === 0) this.syncRevert();
        else this.markDirty();
    }
    redo(): void {
        if (!this.future.length) return;
        this.past.push(this.cloneBlocks());
        this.model.blocks = this.future.pop()!;
        this.renderAll();
        this.updateUndoRedo();
        if (this.past.length === 0) this.syncRevert();
        else this.markDirty();
    }
    private updateUndoRedo(): void {
        if (this.undoBtn) this.undoBtn.disabled = !this.past.length;
        if (this.redoBtn) this.redoBtn.disabled = !this.future.length;
    }

    // ---------- 选中 / 命令 ----------
    private select(index: number): void {
        if (this.current === index) return;
        this.current = index;
        this.docEl?.querySelectorAll('.oe-block').forEach((el) => {
            el.classList.toggle('sel', Number((el as HTMLElement).dataset.idx) === index);
        });
    }

    private exec(cmd: string): void {
        this.restoreSel();
        document.execCommand('styleWithCSS', false, 'true');
        document.execCommand(cmd);
        this.afterFmt();
    }

    private execWith(cmd: string, value: string): void {
        this.restoreSel();
        document.execCommand('styleWithCSS', false, 'true');
        document.execCommand(cmd, false, value);
        this.afterFmt();
    }

    /** 字符级字号（磅）：execCommand fontSize 产生 <font size>，readEditable 再解析回磅值。 */
    private execFontSize(pt: number): void {
        this.restoreSel();
        document.execCommand('styleWithCSS', false, 'false');
        document.execCommand('fontSize', false, ptToN(pt));
        this.afterFmt();
    }

    private afterFmt(): void {
        const ed = this.editingEl();
        const commit = ed && (ed as { __commit?: () => void }).__commit;
        if (commit) commit();
        this.markDirty();
        this.updateFmtState();
    }

    private editingEl(): HTMLElement | null {
        const sel = window.getSelection();
        if (sel && sel.rangeCount) {
            let node = sel.getRangeAt(0).commonAncestorContainer as Node;
            if (node.nodeType === Node.TEXT_NODE) node = node.parentNode as Node;
            const ed = (node as HTMLElement).closest?.('.oe-edit') as HTMLElement | null;
            if (ed) return ed;
        }
        const active = document.activeElement as HTMLElement | null;
        return active && active.classList.contains('oe-edit') ? active : null;
    }

    private saveSel(): void {
        const sel = window.getSelection();
        if (!sel || !sel.rangeCount) return;
        const range = sel.getRangeAt(0);
        let node: Node = range.commonAncestorContainer;
        if (node.nodeType === Node.TEXT_NODE) node = node.parentNode as Node;
        if ((node as HTMLElement).closest?.('.oe-edit')) this.savedRange = range.cloneRange();
    }

    private restoreSel(): void {
        if (!this.savedRange) return;
        const sel = window.getSelection();
        if (sel) {
            sel.removeAllRanges();
            sel.addRange(this.savedRange);
        }
    }

    private updateFmtState(): void {
        const map: Record<string, string> = { bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strikeThrough' };
        for (const [key, cmd] of Object.entries(map)) {
            const btn = this.fmtBtns[key];
            if (!btn) continue;
            let on = false;
            try {
                on = document.queryCommandState(cmd);
            } catch {
                on = false;
            }
            btn.classList.toggle('active', on);
        }
    }

    private currentBlock(): DocxBlock | null {
        return this.current >= 0 ? this.model.blocks[this.current] : null;
    }

    private setBlockType(value: string): void {
        const b = this.currentBlock();
        if (!b) return;
        this.checkpoint();
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
        this.checkpoint();
        (b as DocxParagraph).align = a as DocxParagraph['align'];
        // 局部更新样式，避免重建丢失焦点
        const el = this.docEl?.querySelector(`.oe-block[data-idx="${this.current}"] .oe-edit`) as HTMLElement | null;
        if (el) el.style.textAlign = a === 'left' ? '' : a;
        this.markDirty();
    }

    private insertBlock(block: DocxBlock, after = this.current): void {
        this.checkpoint();
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

    /** 选择本地图片，读为 data URL 后作为 image 块插入（尺寸按 pt 归一）。 */
    private pickImage(): void {
        const input = h('input', { type: 'file', accept: 'image/*', style: 'display:none' }) as HTMLInputElement;
        input.addEventListener('change', () => {
            const file = input.files && input.files[0];
            if (!file) { input.remove(); return; }
            const reader = new FileReader();
            reader.onload = () => {
                const src = String(reader.result || '');
                const img = new Image();
                img.onload = () => {
                    const maxW = 480;
                    let w = img.naturalWidth;
                    let ht = img.naturalHeight;
                    if (w > maxW) { ht = Math.round((ht * maxW) / w); w = maxW; }
                    this.insertBlock({ type: 'image', src, width: Math.round(w * 0.75), height: Math.round(ht * 0.75) });
                };
                img.onerror = () => this.insertBlock({ type: 'image', src });
                img.src = src;
                input.remove();
            };
            reader.onerror = () => input.remove();
            reader.readAsDataURL(file);
        });
        document.body.appendChild(input);
        input.click();
    }

    private moveBlock(dir: number): void {
        const i = this.current;
        const j = i + dir;
        if (i < 0 || j < 0 || j >= this.model.blocks.length) return;
        this.checkpoint();
        const [b] = this.model.blocks.splice(i, 1);
        this.model.blocks.splice(j, 0, b);
        this.current = j;
        this.renderAll();
        this.markDirty();
    }

    private deleteBlock(): void {
        if (this.current < 0) return;
        this.checkpoint();
        this.model.blocks.splice(this.current, 1);
        this.current = -1;
        this.renderAll();
        this.markDirty();
    }

    private focusBlock(index: number): void {
        window.setTimeout(() => {
            const ed = this.docEl?.querySelector(`.oe-block[data-idx="${index}"] .oe-edit`) as HTMLElement | null;
            ed?.focus();
        }, 0);
    }

    // ---------- 表格结构 ----------
    private tableAddRow(table: DocxTable, index: number): void {
        this.checkpoint();
        const cols = table.rows[0]?.length || 1;
        (table.rows as DocxCell[][]).push(new Array(cols).fill(''));
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
    private tableDelRow(table: DocxTable, index: number): void {
        if (table.rows.length <= 1) return;
        this.checkpoint();
        (table.rows as DocxCell[][]).pop();
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
    private tableAddCol(table: DocxTable, index: number): void {
        this.checkpoint();
        for (const row of table.rows) (row as DocxCell[]).push('');
        this.current = index;
        this.renderAll();
        this.markDirty();
    }
    private tableDelCol(table: DocxTable, index: number): void {
        const cols = table.rows[0]?.length || 0;
        if (cols <= 1) return;
        this.checkpoint();
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
