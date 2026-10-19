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
    /** 页边距（pt），默认 56（约 2cm） */
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

/** 任意块 -> 纯文本（用于表格单元格降级） */
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
        case 'list':
            return block.items.map((i) => (Array.isArray(i) ? runsToText(i as DocxRun[]) : '')).join(' ');
        default:
            return '';
    }
}

/** 列表 -> 一组 runs 元素（含项目符号 / 编号与嵌套缩进） */
function listToElements(list: Extract<DocxBlock, { type: 'list' }>, depth: number): PdfElement[] {
    const els: PdfElement[] = [];
    const indent = '   '.repeat(depth);
    list.items.forEach((item, index) => {
        const runs = (Array.isArray(item) ? item : []) as DocxRun[];
        const children = (runs as DocxRun[] & { children?: typeof list.items }).children;
        const prefix = list.ordered ? `${index + 1}. ` : '• ';
        const prefixRun: PdfRun = { text: indent + prefix };
        const itemRuns = runs.map(runToPdf);
        els.push({ type: 'runs', runs: [prefixRun, ...itemRuns], spacingAfter: 2 });
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
        rows: table.rows.map((row) => row.map(cellToText)) as string[][]
    };
}

/** 单个块 -> PDF 元素（可能产出多个，如列表） */
function blockToElements(block: DocxBlock, contentWidth: number): PdfElement[] {
    switch (block.type) {
        case 'heading':
            return [
                {
                    type: 'heading',
                    level: block.level,
                    text: block.runs ? runsToText(block.runs) : block.text ?? ''
                }
            ];
        case 'paragraph':
            return [{ type: 'runs', runs: (block.runs || [{ text: block.text ?? '' }]).map(runToPdf), spacingAfter: 4 }];
        case 'list':
            return listToElements(block, 0);
        case 'table':
            return [tableToElement(block)];
        case 'image': {
            const src =
                typeof block.src === 'string'
                    ? block.src
                    : block.assetId
                      ? ({ assetId: block.assetId } as never)
                      : (block.src as never);
            return [{ type: 'image', src, width: block.width, height: block.height }];
        }
        case 'code':
            return [{ type: 'runs', runs: [{ text: block.text, code: true }], spacingAfter: 4 }];
        case 'quote':
            return [
                {
                    type: 'runs',
                    runs: [{ text: block.runs ? runsToText(block.runs) : block.text ?? '' }],
                    spacingAfter: 4
                }
            ];
        case 'hr':
            return [{ type: 'line', x1: 0, y1: 0, x2: contentWidth, y2: 0, lineWidth: 0.5 }];
        case 'pageBreak':
            return [{ type: 'pageBreak' }];
        case 'textBox':
            return block.paragraphs.flatMap((p) => blockToElements(p, contentWidth));
        case 'toc':
            return [];
        default:
            return [];
    }
}

/** 内容区宽度（pt） */
function contentWidthOf(options: BlocksToPdfOptions): number {
    const size: [number, number] =
        options.pageSize && Array.isArray(options.pageSize)
            ? options.pageSize
            : (options.pageSize as string) === 'A3'
              ? [841.89, 1190.55]
              : (options.pageSize as string) === 'A5'
                ? [419.53, 595.28]
                : (options.pageSize as string) === 'Letter'
                  ? [612, 792]
                  : (options.pageSize as string) === 'Legal'
                    ? [612, 1008]
                    : [595.28, 841.89];
    const m = typeof options.margin === 'number' ? options.margin : 56;
    const mx = typeof options.margin === 'object' ? Math.max(options.margin.left || m, options.margin.right || m) : m;
    return Math.max(100, size[0] - mx * 2);
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
    const contentWidth = contentWidthOf(options);

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

    const elements: PdfElement[] = blocks.flatMap((b) => blockToElements(b, contentWidth));

    const spec: PdfDocumentSpec = {
        metadata: Array.isArray(input) ? undefined : input.metadata,
        elements,
        fonts,
        assets: resolveAssets(options.assets)
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
