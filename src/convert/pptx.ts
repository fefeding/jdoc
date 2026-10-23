/**
 * PPTX -> 文档块模型（DocxBlock）
 *
 * PPTX 的内容模型是「幻灯片 + 绝对定位元素」，与 docx/md/html 的块序列不同，
 * 这里把每张幻灯片映射成「标题 + 正文块」：
 *
 *   幻灯片标题（字号最大的文本框）-> 标题（heading）
 *   其余文本框（含项目符号）      -> 段落 / 列表
 *   表格元素                      -> 表格（保留合并）
 *   图片元素                      -> 图片
 *   图表元素                      -> 标题 + 数据表
 *   SmartArt / 图示              -> 列表
 *
 * 这样 pptx -> md / html / docx / pdf 四条目标路径都能拿到接近原稿的内容。
 *
 * @module convert/pptx
 */

import type { DocInput } from '../core/input';
import type { DocAsset, DocAssetMode } from '../core/asset';
import { pptxToStandard } from '../pptx/index';
// PPTX 元素级类型由底层 @fefeding/ppt-parser 提供（src/pptx/index 只再导出 PptxDocument 等少数类型）
import type {
    PptxDocument,
    PptxSlide,
    PptxElement,
    PptxTextElement,
    PptxParagraph,
    PptxTextRun,
    PptxTableCell,
    PptxTableElement,
    PptxImageElement,
    PptxChartElement,
    PptxDiagramElement,
    PptxGroupElement,
    PptxMathElement,
    PptxGaugeElement,
    PptxDiagramNode,
    PptxMetadata
} from '@fefeding/ppt-parser';
import { pptxChartToSpec } from '../pptx/chart';
import type { DocxBlock, DocxCell, DocxCellDetail, DocxHeading, DocxMetadata, DocxRun, DocxTable } from '../docx/types';

/** pptxToBlocks 选项 */
export interface PptxToBlocksOptions {
    /** 是否把每张幻灯片输出一个标题，默认 true */
    slideHeading?: boolean;
    /** 幻灯片标题使用的标题级别，默认 1 */
    headingLevel?: number;
    /** 是否把演讲者备注作为引用块输出，默认 false */
    notes?: boolean;
    /** 图片资源处理方式，默认 'both'（元素上保留 data URL，同时抽出资源表） */
    assets?: DocAssetMode;
}

/** pptxToBlocks 结果 */
export interface PptxToBlocksResult {
    blocks: DocxBlock[];
    metadata?: DocxMetadata;
    assets?: DocAsset[];
}

/** 一个可参与「列表分组」的段落项 */
interface ParaItem {
    runs: DocxRun[];
    bullet: false | true | 'number';
    align?: string;
}

/** PPTX 解析出的元素 -> 段落项 / 独立块 */
type ElementItem = { kind: 'para'; item: ParaItem } | { kind: 'block'; block: DocxBlock };

// ---- runs 构造 ----

/** PPTX 文本片段 -> DocxRun */
function runToRun(run: PptxTextRun): DocxRun {
    const r: DocxRun = { text: run.text ?? '' };
    if (run.bold) r.bold = true;
    if (run.italic) r.italic = true;
    if (run.underline) r.underline = true;
    if (run.strike) r.strike = true;
    if (run.color) r.color = run.color;
    if (run.fontFaceEa || run.fontFace) r.font = run.fontFaceEa || run.fontFace;
    if (run.fontSize) r.size = run.fontSize;
    if (run.href) r.link = run.href;
    return r;
}

/** PPTX 段落 -> runs（富文本优先，否则取纯文本） */
function paragraphToRuns(p: PptxParagraph): DocxRun[] {
    if (p.runs?.length) return p.runs.map(runToRun);
    if (p.text != null) return [{ text: p.text }];
    return [];
}

/** 段落是否为项目符号 / 编号 */
function paragraphBullet(p: PptxParagraph): false | true | 'number' {
    if (!p.bullet) return false;
    if (p.bullet === true) return true;
    if (p.bullet === 'number') return 'number';
    if (typeof p.bullet === 'object') return p.bullet.type === 'number' ? 'number' : true;
    return true;
}

// ---- 单元格 / 表格 ----

/** 表格单元格 -> DocxCell（合并的延续单元格返回 null 表示跳过） */
function pptxCellToBlockCell(cell: PptxTableCell): DocxCell | null {
    // PPTX 合并：主单元格带 colSpan/rowSpan，后续单元格 hMerge/vMerge=true 仅占位
    if (cell.hMerge || cell.vMerge) return null;
    const runs = cell.paragraphs?.length ? cell.paragraphs.flatMap(paragraphToRuns) : cell.text ? [{ text: cell.text }] : [];
    const text = cell.text ?? cell.paragraphs?.map((p) => p.text ?? '').join('') ?? '';
    const detail: DocxCellDetail = { type: 'cell', text };
    if (runs.length) detail.runs = runs;
    if (cell.colSpan && cell.colSpan > 1) detail.colspan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1) detail.rowspan = cell.rowSpan;
    if (cell.fill) detail.fill = cell.fill;
    return detail;
}

/** PPTX 表格 -> DocxBlock 表格 */
function pptxTableToBlock(table: PptxTableElement): DocxBlock {
    const rows: DocxCell[][] = [];
    for (const row of table.rows) {
        const cells: DocxCell[] = [];
        for (const cell of row.cells) {
            const blockCell = pptxCellToBlockCell(cell);
            if (blockCell) cells.push(blockCell);
        }
        rows.push(cells);
    }
    return { type: 'table', header: rows.length > 1, rows };
}

// ---- 图表 / 图示 / 数学 / 仪表 ----

/** 构造一个纯文本单元格（避免字面量 type 在数组上下文被拓宽为 string） */
function textCell(text: string): DocxCellDetail {
    return { type: 'cell', text };
}

/**
 * 图表 -> 图表块
 *
 * HTML 目标会把 spec 渲染成内嵌矢量 SVG；md / docx / pdf 画不了矢量图，
 * 退化成随块携带的兜底数据表。
 */
function pptxChartToBlocks(chart: PptxChartElement): DocxBlock[] {
    const series = chart.series ?? [];
    const categories = chart.categories ?? [];
    const header: DocxCell[] = [textCell(''), ...series.map((s) => textCell(s.name ?? ''))];
    const rows: DocxCell[][] = categories.map((cat, i) => [
        textCell(cat),
        ...series.map((s) => textCell(String(s.values?.[i] ?? '')))
    ]);
    const fallback: DocxTable = { type: 'table', header: true, rows: [header, ...rows] };
    return [
        {
            type: 'chart',
            spec: pptxChartToSpec(chart as unknown as Parameters<typeof pptxChartToSpec>[0]),
            width: Number(chart.width) || undefined,
            height: Number(chart.height) || undefined,
            caption: chart.title,
            fallback
        }
    ];
}

/** 图示节点展平为文本数组 */
function flattenDiagram(nodes?: PptxDiagramNode[]): string[] {
    const out: string[] = [];
    const walk = (list?: PptxDiagramNode[]) => {
        for (const node of list || []) {
            if (node.text) out.push(node.text);
            walk(node.children);
        }
    };
    walk(nodes);
    return out;
}

// ---- 文本框 / 单元素 -> 段落项 ----

/** 取文本框的首段可见文本（用作幻灯片标题） */
function firstTextOf(el: PptxTextElement): string {
    for (const p of el.paragraphs ?? []) {
        const t = p.text ?? p.runs?.map((r) => r.text).join('') ?? '';
        if (t.trim()) return t.trim();
    }
    if (el.text?.trim()) return el.text.trim();
    if (el.runs?.length) {
        const t = el.runs.map((r) => r.text).join('').trim();
        if (t) return t;
    }
    return '';
}

/** 文本框内最大的字号（用于挑选幻灯片标题） */
function maxFontSizeOf(el: PptxTextElement): number {
    let max = el.fontSize ?? 0;
    for (const p of el.paragraphs ?? []) {
        for (const r of p.runs ?? []) max = Math.max(max, r.fontSize ?? 0);
    }
    return max;
}

/** 文本框 -> 段落项（每个段落一项，保留项目符号） */
function textElementToItems(el: PptxTextElement): ParaItem[] {
    const items: ParaItem[] = [];
    const paragraphs = el.paragraphs ?? [];
    if (paragraphs.length) {
        for (const p of paragraphs) {
            const runs = paragraphToRuns(p);
            if (!runs.some((r) => r.text)) continue;
            items.push({ runs, bullet: paragraphBullet(p), align: p.align });
        }
    }
    if (!items.length) {
        if (el.runs?.length) {
            const runs = el.runs.map(runToRun);
            if (runs.some((r) => r.text)) items.push({ runs, bullet: false });
        } else if (el.text) {
            items.push({ runs: [{ text: el.text }], bullet: false });
        }
    }
    return items;
}

/** 单个元素 -> 段落项 / 独立块 */
function elementToItems(el: PptxElement): ElementItem[] {
    switch (el.type) {
        case 'text':
            return textElementToItems(el as PptxTextElement).map((item) => ({ kind: 'para', item }));
        case 'table':
            return [{ kind: 'block', block: pptxTableToBlock(el as PptxTableElement) }];
        case 'image': {
            const img = el as PptxImageElement;
            const src = img.data || img.src;
            if (!src) return [];
            const block: DocxBlock = {
                type: 'image',
                src,
                width: img.width,
                height: img.height,
                caption: img.name
            };
            return [{ kind: 'block', block }];
        }
        case 'chart':
            return pptxChartToBlocks(el as PptxChartElement).map((block) => ({ kind: 'block', block }));
        case 'diagram': {
            const diagram = el as PptxDiagramElement;
            const texts = diagram.texts?.length ? diagram.texts : flattenDiagram(diagram.nodes);
            return texts
                .filter((t) => t && t.trim())
                .map((t) => ({ kind: 'para', item: { runs: [{ text: t }], bullet: true } }));
        }
        case 'group': {
            const out: ElementItem[] = [];
            for (const child of (el as PptxGroupElement).children ?? []) out.push(...elementToItems(child));
            return out;
        }
        case 'math': {
            const t = (el as PptxMathElement).text?.trim();
            return t ? [{ kind: 'para', item: { runs: [{ text: t }], bullet: false } }] : [];
        }
        case 'gauge': {
            const g = el as PptxGaugeElement;
            const label = g.label ?? '值';
            return [{ kind: 'para', item: { runs: [{ text: `${label}: ${g.value}` }], bullet: false } }];
        }
        // shape / video / audio / ole / connector / raw 等图形/媒体：转换场景下无文本可提取，跳过
        default:
            return [];
    }
}

/** 把段落项流分组为段落块与列表块 */
function groupItems(items: ParaItem[]): DocxBlock[] {
    const blocks: DocxBlock[] = [];
    let i = 0;
    while (i < items.length) {
        const item = items[i];
        if (item.bullet) {
            const ordered = item.bullet === 'number';
            const listItems: ParaItem[] = [];
            while (i < items.length && items[i].bullet) {
                listItems.push(items[i]);
                i++;
            }
            blocks.push({
                type: 'list',
                ordered,
                items: listItems.map((li) => li.runs)
            });
        } else {
            blocks.push({ type: 'paragraph', align: item.align as never, runs: item.runs });
            i++;
        }
    }
    return blocks;
}

/** 单张幻灯片 -> 块 */
function slideToBlocks(
    slide: PptxSlide,
    index: number,
    options: { slideHeading: boolean; headingLevel: number; notes: boolean }
): DocxBlock[] {
    const blocks: DocxBlock[] = [];
    const elements = slide.elements ?? [];

    // 1) 挑选标题：字号最大的文本框
    const textEls = elements.filter((e): e is PptxTextElement => e.type === 'text');
    let titleEl: PptxTextElement | null = null;
    let best = -1;
    for (const el of textEls) {
        const fs = maxFontSizeOf(el);
        if (fs > best) {
            best = fs;
            titleEl = el;
        }
    }
    const titleText = titleEl ? firstTextOf(titleEl) : '';

    if (options.slideHeading) {
        blocks.push({
            type: 'heading',
            level: options.headingLevel as DocxHeading['level'],
            text: titleText || `幻灯片 ${index + 1}`
        });
    }

    // 2) 正文：排除标题文本框，按从上到下、从左到右排序
    const bodyEls = elements.filter((e) => e !== titleEl);
    bodyEls.sort((a, b) => (a.y ?? 0) - (b.y ?? 0) || (a.x ?? 0) - (b.x ?? 0));

    const items: ParaItem[] = [];
    const direct: DocxBlock[] = [];
    for (const el of bodyEls) {
        const elementItems = elementToItems(el);
        for (const ei of elementItems) {
            if (ei.kind === 'para') items.push(ei.item);
            else direct.push(ei.block);
        }
    }
    blocks.push(...groupItems(items));
    blocks.push(...direct);

    // 3) 演讲者备注（可选）
    if (options.notes && slide.notes?.trim()) {
        blocks.push({ type: 'quote', runs: [{ text: slide.notes.trim() }] });
    }
    return blocks;
}

/** PPTX 元数据 -> 文档元数据 */
function toDocxMetadata(metadata?: PptxMetadata): DocxMetadata | undefined {
    if (!metadata) return undefined;
    const result: DocxMetadata = {};
    const put = (key: keyof DocxMetadata, value: unknown) => {
        if (value === undefined || value === null || value === '') return;
        result[key] = String(value);
    };
    put('title', metadata.title);
    put('subject', metadata.subject);
    put('author', metadata.author);
    put('keywords', metadata.keywords);
    put('description', metadata.description);
    put('lastModifiedBy', metadata.lastModifiedBy);
    return Object.keys(result).length ? result : undefined;
}

/**
 * 把已解析的 PPTX 标准文档映射为文档块模型
 *
 * @param doc - pptxToStandard 的解析结果
 * @param options - 标题级别、备注等选项
 */
export function pptxDocumentToBlocks(
    doc: PptxDocument,
    options: PptxToBlocksOptions = {}
): PptxToBlocksResult {
    const { slideHeading = true, headingLevel = 1, notes = false } = options;
    const blocks: DocxBlock[] = [];
    for (let i = 0; i < doc.slides.length; i++) {
        blocks.push(...slideToBlocks(doc.slides[i], i, { slideHeading, headingLevel, notes }));
    }
    return {
        blocks,
        metadata: toDocxMetadata(doc.metadata)
    };
}

/**
 * 把 PPTX 解析为文档块模型
 *
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 标题、备注、图片资源等选项
 * @returns 块序列与元数据
 *
 * @example
 * ```js
 * const { blocks } = await pptxToBlocks(file);
 * console.log(blocks[0]); // { type: 'heading', level: 1, text: '幻灯片标题' }
 * ```
 */
export async function pptxToBlocks(
    input: DocInput,
    options: PptxToBlocksOptions = {}
): Promise<PptxToBlocksResult> {
    const { assets = 'both', ...rest } = options;
    const doc = (await pptxToStandard(input, { assets })) as PptxDocument;
    const result = pptxDocumentToBlocks(doc, rest);
    const parsedAssets = (doc as { assets?: DocAsset[] }).assets;
    if (parsedAssets?.length) result.assets = parsedAssets;
    return result;
}
