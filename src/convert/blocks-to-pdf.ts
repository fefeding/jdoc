/**
 * 文档块模型（DocxBlock）-> PDF
 *
 * 把统一 IR 映射为 `jsonToPdf` 所需的 `PdfDocumentSpec`：
 * 标题/段落/列表/表格/图片/分隔线都能映射，段落与列表保留行内样式（加粗/斜体/
 * 彩色/超链接/行内代码）。遇到中文且未显式提供字体时，自动尝试探测系统字体，
 * 让 Node 端「开箱即用」地输出中文 PDF。
 *
 * @module convert/blocks-to-pdf
 */

import { jsonToPdf } from '../pdf/index';
import { detectSystemFonts } from '../pdf/fonts';
import { getAsset, type DocAsset } from '../core/asset';
import type { PdfDocumentSpec, PdfElement, PdfFontSet, PdfPageSize, PdfRun, PdfTableElement } from '../pdf/types';
import type { DocxBlock, DocxCell, DocxDocument, DocxRun } from '../docx/types';

/** blocksToPdf 选项 */
export interface BlocksToPdfOptions {
    /** 内嵌字体（中文必须）；未提供且内容含中文时自动探测系统字体 */
    fonts?: PdfFontSet;
    /** 资源表，用于解析图片块的 assetId 引用 */
    assets?: DocAsset[] | Record<string, DocAsset> | { assets?: DocAsset[] };
    /** 纸张尺寸，默认 A4 */
    pageSize?: PdfPageSize;
    /** 页边距（pt），默认 Word 中文标准：上下 72（2.54cm）、左右 90（3.17cm） */
    margin?: number | { top?: number; right?: number; bottom?: number; left?: number };
    /** 是否自动探测系统字体（默认 true，仅 Node 端有效） */
    autoFont?: boolean;
}

/** 行内 run -> PDF run */
function runToPdf(r: DocxRun): PdfRun {
    return {
        text: r.text,
        bold: r.bold,
        italic: r.italic,
        strike: r.strike,
        color: r.color,
        link: r.link,
        code: r.code,
        size: r.size
    };
}

/** run 序列 -> 纯文本 */
function runsToText(runs: DocxRun[] | undefined): string {
    return (runs || []).map((r) => r.text).join('');
}

/**
 * Word/WPS 中文排版基线（对齐目标：与 WPS 打开同一 docx 的分页位置接近）
 *
 * - 正文五号 10.5pt；A4 纵向页边距上下 2.54cm(72pt)、左右 3.17cm(90pt)
 * - 单倍行距：中文约 1.3 倍字号（宋体系），西文约 1.15 倍
 * - 段前/段后默认 0（Word 默认段间距为 0，间距来自行距）
 */
const WORD_DEFAULTS = {
    /** 正文默认字号（pt，五号） */
    bodySize: 10.5,
    /** A4 页边距（pt）：上下 2.54cm、左右 3.17cm（Word 中文默认） */
    margin: { top: 72, bottom: 72, left: 90, right: 90 },
    /** 中文单倍行距系数 */
    lineBaseCjk: 1.3,
    /** 西文单倍行距系数 */
    lineBaseLatin: 1.15
};

/** Word 中文内置标题字号（pt）：h1 二号 22、h2 三号 16、h3 小三 15、h4 四号 14、h5 小四 12、h6 五号 10.5 */
const HEADING_SIZES = [22, 16, 15, 14, 12, 10.5];

/** Word 中文内置标题段前/段后间距（pt） */
const HEADING_SPACING: [number, number][] = [
    [17, 16],
    [13, 13],
    [13, 13],
    [12, 10],
    [12, 10],
    [12, 10]
];

/** 行距倍数（1 = 单倍）-> PDF 行高系数；缺省按内容 CJK 与否取单倍行距基线 */
function lineSpacingToLineHeight(lineSpacing: number | undefined, hasCjk: boolean): number {
    const base = hasCjk ? WORD_DEFAULTS.lineBaseCjk : WORD_DEFAULTS.lineBaseLatin;
    return lineSpacing && lineSpacing > 0 ? base * lineSpacing : base;
}

/** 单元格 -> 纯文本 */
function cellToText(cell: DocxCell | undefined): string {
    if (cell == null) return '';
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return runsToText(cell);
    const detail = cell as { blocks?: DocxBlock[]; runs?: DocxRun[]; text?: string };
    if (detail.blocks) return detail.blocks.map(blockToText).join(' ');
    if (detail.runs) return runsToText(detail.runs);
    return detail.text ?? '';
}

/** 任意块 -> 纯文本（用于表格单元格降级与 CJK 检测） */
function blockToText(block: DocxBlock): string {
    switch (block.type) {
        case 'heading':
            return block.runs ? runsToText(block.runs) : block.text ?? '';
        case 'paragraph':
            return block.runs ? runsToText(block.runs) : block.text ?? '';
        case 'code':
            return block.text;
        case 'image':
            return block.alt ?? '';
        case 'textBox':
            return block.paragraphs.map(blockToText).join(' ');
        case 'quote':
            return block.runs ? runsToText(block.runs) : block.text ?? '';
        case 'list':
            return block.items
                .map((i) => (typeof i === 'string' ? i : Array.isArray(i) ? runsToText(i as DocxRun[]) : ''))
                .join(' ');
        case 'table':
            return block.rows.map((row) => row.map(cellToText).join(' ')).join(' ');
        default:
            return '';
    }
}

/** 列表 -> 一组 runs 元素（含项目符号 / 编号与嵌套缩进） */
function listToElements(list: Extract<DocxBlock, { type: 'list' }>, depth: number): PdfElement[] {
    const els: PdfElement[] = [];
    const indent = '   '.repeat(depth);
    list.items.forEach((item, index) => {
        const runs = (typeof item === 'string' ? [{ text: item }] : Array.isArray(item) ? item : []) as DocxRun[];
        const children = (runs as DocxRun[] & { children?: typeof list.items }).children;
        const prefix = list.ordered ? `${index + 1}. ` : '• ';
        const prefixRun: PdfRun = { text: indent + prefix };
        const itemRuns = runs.map(runToPdf);
        const text = runsToText(runs);
        // 列表项行距与正文一致、段间距 0（Word 列表即普通段落）
        els.push({
            type: 'runs',
            runs: [prefixRun, ...itemRuns],
            spacingAfter: 0,
            lineHeight: lineSpacingToLineHeight(undefined, hasCjkText(text))
        });
        if (children?.length) {
            els.push(...listToElements({ type: 'list', ordered: list.ordered, items: children }, depth + 1));
        }
    });
    return els;
}

/** 表格 -> PDF 表格元素 */
function tableToElement(table: Extract<DocxBlock, { type: 'table' }>): PdfTableElement {
    return {
        type: 'table',
        header: table.header,
        rows: table.rows.map((row) => row.map(cellToText)) as string[][],
        // 透传 docx tblGrid 列宽（磅），缺列宽时由生成端按内容区等分
        columnWidths: table.columnWidths,
        fontSize: WORD_DEFAULTS.bodySize
    };
}

/** 单个块 -> PDF 元素（可能产出多个，如列表） */
function blockToElements(block: DocxBlock): PdfElement[] {
    switch (block.type) {
        case 'heading': {
            const idx = Math.min(5, Math.max(0, block.level - 1));
            return [
                {
                    type: 'heading',
                    level: block.level,
                    text: block.runs ? runsToText(block.runs) : block.text ?? '',
                    // Word 中文内置标题字号与段前/段后间距
                    size: HEADING_SIZES[idx],
                    spacingBefore: HEADING_SPACING[idx][0],
                    spacingAfter: HEADING_SPACING[idx][1],
                    lineHeight: lineSpacingToLineHeight(undefined, true)
                }
            ];
        }
        case 'paragraph': {
            const runs = (block.runs || [{ text: block.text ?? '' }]).map(runToPdf);
            const text = runsToText(block.runs) || block.text || '';
            // justify 的断行/分页与 left 完全一致（断行只取决于宽度），暂按左对齐渲染
            const align = block.align === 'justify' ? undefined : block.align;
            return [
                {
                    type: 'runs',
                    runs,
                    align,
                    indentLeft: block.indentLeft,
                    indentRight: block.indentRight,
                    indentFirstLine: block.indentFirstLine,
                    // Word 默认段前/段后均为 0（间距来自行距）
                    spacingBefore: block.spacingBefore ?? 0,
                    spacingAfter: block.spacingAfter ?? 0,
                    lineHeight: lineSpacingToLineHeight(block.lineSpacing, hasCjkText(text))
                }
            ];
        }
        case 'list':
            return listToElements(block, 0);
        case 'table':
            return [tableToElement(block)];
        case 'image': {
            const src =
                typeof block.src === 'string'
                    ? block.src
                    : block.assetId
                      ? { assetId: block.assetId }
                      : block.src;
            return [{ type: 'image', src, width: block.width, height: block.height }];
        }
        case 'code':
            return [{ type: 'runs', runs: [{ text: block.text, code: true }], spacingAfter: 4 }];
        case 'quote': {
            const text = block.runs ? runsToText(block.runs) : block.text ?? '';
            return [
                {
                    type: 'runs',
                    runs: [{ text }],
                    // Word 引用样式：左缩进 1cm
                    indentLeft: 28.35,
                    spacingAfter: 4,
                    lineHeight: lineSpacingToLineHeight(undefined, hasCjkText(text))
                }
            ];
        }
        case 'hr':
            // 流式分隔线：画在当前内容位置（line 元素是绝对定位语义，不适合流式内容）
            return [{ type: 'hr', lineWidth: 0.5 }];
        case 'pageBreak':
            return [{ type: 'pageBreak' }];
        case 'textBox':
            return block.paragraphs.flatMap((p) => blockToElements(p));
        case 'toc':
            return [];
        default:
            return [];
    }
}

/** 文本是否含 CJK 字符（决定单倍行距基线） */
function hasCjkText(text: string): boolean {
    return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF01-\uFF5E]/.test(text);
}

/**
 * 把文档块模型渲染为 PDF 字节
 * @param input - 块序列或完整文档（可带元数据）
 * @param options - 字体 / 资源 / 页面选项
 * @returns PDF 字节数据
 */
export async function blocksToPdf(
    input: DocxBlock[] | DocxDocument,
    options: BlocksToPdfOptions = {}
): Promise<Uint8Array> {
    const blocks = Array.isArray(input) ? input : input.blocks || [];

    // 自动字体：含中文且未提供字体时探测系统字体（仅 Node 端）
    let fonts = options.fonts;
    if (!fonts && options.autoFont !== false) {
        const hasCjk = blocks.some((b) => blockHasCjk(b));
        if (hasCjk) {
            try {
                fonts = await detectSystemFonts();
            } catch {
                // 浏览器端无系统字体：交给 jsonToPdf 抛友好的中文提示
            }
        }
    }

    const elements: PdfElement[] = blocks.flatMap((b) => blockToElements(b));

    const spec: PdfDocumentSpec = {
        metadata: Array.isArray(input) ? undefined : input.metadata,
        elements,
        fonts,
        assets: resolveAssets(options.assets),
        // Word 中文排版基线：正文五号 10.5pt、A4 边距上下 72pt / 左右 90pt
        defaultSize: WORD_DEFAULTS.bodySize,
        pageSize: options.pageSize ?? 'A4',
        margin: options.margin ?? WORD_DEFAULTS.margin
    };

    return jsonToPdf(spec, { assets: resolveAssets(options.assets) });
}

/** 资源表归一 */
function resolveAssets(assets: BlocksToPdfOptions['assets']): DocAsset[] | Record<string, DocAsset> | undefined {
    if (!assets) return undefined;
    if (Array.isArray(assets)) return assets;
    const explicit = (assets as { assets?: DocAsset[] }).assets;
    if (explicit) return explicit;
    return assets as Record<string, DocAsset>;
}

/** 块是否含中文 */
function blockHasCjk(block: DocxBlock): boolean {
    const text = blockToText(block);
    return /[一-鿿豈-﫿＀-￯]/.test(text);
}
