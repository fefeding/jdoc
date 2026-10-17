/**
 * PDF 生成
 *
 * 基于 pdf-lib 输出 PDF。坐标系对调用方保持「顶部原点」（与 HTML 心智模型一致），
 * 内部换算为 pdf-lib 的左下原点。
 *
 * 中文字体：pdf-lib 的 14 种标准字体不含 CJK 字形，需要通过 fonts 传入 TTF/OTF 字节；
 * 未传入时使用标准字体，此时中文会抛出明确错误而不是静默输出空白。
 *
 * @module pdf/generate
 */

import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFImage, type RGB } from 'pdf-lib';
import { JdocError, wrapError } from '../core/errors';
import { toUint8Array } from '../core/input';
import type {
    PdfDocumentSpec,
    PdfElement,
    PdfFontSet,
    PdfGenerateOptions,
    PdfPageSize,
    PdfTableElement,
    PdfTextElement
} from './types';

/** 纸张尺寸（pt，1pt = 1/72 英寸） */
const PAGE_SIZES: Record<string, [number, number]> = {
    A3: [841.89, 1190.55],
    A4: [595.28, 841.89],
    A5: [419.53, 595.28],
    Letter: [612, 792],
    Legal: [612, 1008]
};

/** 命名颜色 */
const NAMED_COLORS: Record<string, string> = {
    black: '#000000',
    white: '#FFFFFF',
    red: '#FF0000',
    green: '#008000',
    blue: '#0000FF',
    gray: '#808080',
    grey: '#808080',
    silver: '#C0C0C0',
    yellow: '#FFFF00',
    orange: '#FFA500',
    purple: '#800080',
    navy: '#000080',
    teal: '#008080',
    maroon: '#800000',
    olive: '#808000',
    lime: '#00FF00',
    aqua: '#00FFFF',
    cyan: '#00FFFF',
    fuchsia: '#FF00FF',
    magenta: '#FF00FF'
};

/** 颜色字符串 -> RGB；无法解析时返回黑色 */
export function parseColor(value: string | undefined, fallback: RGB = rgb(0, 0, 0)): RGB {
    if (!value) return fallback;
    let hex = value.trim();

    const named = NAMED_COLORS[hex.toLowerCase()];
    if (named) hex = named;
    if (hex.startsWith('#')) hex = hex.slice(1);
    if (hex.length === 3) hex = hex.split('').map((ch) => ch + ch).join('');
    if (!/^[0-9a-fA-F]{6}$/.test(hex)) return fallback;

    return rgb(
        parseInt(hex.slice(0, 2), 16) / 255,
        parseInt(hex.slice(2, 4), 16) / 255,
        parseInt(hex.slice(4, 6), 16) / 255
    );
}

/** 纸张尺寸 -> [宽, 高]（已考虑横竖屏） */
export function resolvePageSize(
    size: PdfPageSize | undefined,
    landscape: boolean
): [number, number] {
    let width: number;
    let height: number;

    if (Array.isArray(size)) {
        [width, height] = size;
    } else if (size && typeof size === 'object') {
        width = size.width;
        height = size.height;
    } else {
        const key = typeof size === 'string' ? size : 'A4';
        [width, height] = PAGE_SIZES[key] || PAGE_SIZES.A4;
    }

    return landscape ? [height, width] : [width, height];
}

/** 页边距归一化 */
function resolveMargin(margin: PdfPageSpecLike['margin']): { top: number; right: number; bottom: number; left: number } {
    const DEFAULT = 56; // 约 2cm
    if (typeof margin === 'number') {
        return { top: margin, right: margin, bottom: margin, left: margin };
    }
    if (margin && typeof margin === 'object') {
        return {
            top: margin.top ?? DEFAULT,
            right: margin.right ?? DEFAULT,
            bottom: margin.bottom ?? DEFAULT,
            left: margin.left ?? DEFAULT
        };
    }
    return { top: DEFAULT, right: DEFAULT, bottom: DEFAULT, left: DEFAULT };
}

/** 页边距类型别名 */
type PdfPageSpecLike = { margin?: number | { top?: number; right?: number; bottom?: number; left?: number } };

/** 字体集合：按 normal/bold/italic/boldItalic 分类 */
interface FontBundle {
    normal: PDFFont;
    bold: PDFFont;
    italic: PDFFont;
    boldItalic: PDFFont;
    /** 是否为自定义内嵌字体（决定中文是否可用） */
    embedded: boolean;
}

/** 标准字体族名 -> pdf-lib 常量 */
const STANDARD_FAMILIES = [
    { key: 'normal', constant: StandardFonts.Helvetica },
    { key: 'bold', constant: StandardFonts.HelveticaBold },
    { key: 'italic', constant: StandardFonts.HelveticaOblique },
    { key: 'boldItalic', constant: StandardFonts.HelveticaBoldOblique }
] as const;

/** 内嵌字体字节归一化 */
async function toFontBytes(source: string | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
    if (typeof source !== 'string') return toUint8Array(source);
    // base64 或 data URL
    const match = /^data:[^;,]*(;base64)?,(.*)$/s.exec(source.trim());
    const payload = match ? match[2] : source;
    const clean = payload.replace(/\s+/g, '');
    const atobFn = (globalThis as { atob?: (s: string) => string }).atob;
    const binary = atobFn ? atobFn(clean) : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i) & 0xff;
    return bytes;
}

/** 构建字体集合 */
async function buildFonts(
    doc: PDFDocument,
    fonts: PdfFontSet | undefined
): Promise<FontBundle> {
    if (fonts && (fonts.regular || fonts.bold)) {
        const embed = async (source: string | ArrayBuffer | Uint8Array | undefined, fallback?: string) => {
            const bytes = source ? await toFontBytes(source) : fallback ? await toFontBytes(fallback) : undefined;
            return bytes ? doc.embedFont(bytes) : undefined;
        };

        const regular = await embed(fonts.regular);
        const normal = regular || (await doc.embedFont(StandardFonts.Helvetica));
        const bold = (await embed(fonts.bold, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const italic = (await embed(fonts.italic, typeof fonts.regular === 'string' ? fonts.regular : undefined)) || normal;
        const boldItalic =
            (await embed(fonts.boldItalic, typeof fonts.bold === 'string' ? fonts.bold : undefined)) || bold;
        return { normal, bold, italic, boldItalic, embedded: true };
    }

    // 标准 14 字体：Helvetica + Times + Courier 共 4 组，取 Helvetica 族做默认
    const normal = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const italic = await doc.embedFont(StandardFonts.HelveticaOblique);
    const boldItalic = await doc.embedFont(StandardFonts.HelveticaBoldOblique);
    return { normal, bold, italic, boldItalic, embedded: false };
}

/** 检测文本中是否含有 CJK 字符（标准字体无法渲染） */
function hasCjk(text: string): boolean {
    return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF]/.test(text);
}

/** 文本按宽度换行（中文按字断行，西文优先按空格断行） */
export function wrapText(text: string, font: PDFFont, size: number, maxWidth: number): string[] {
    if (!text) return [];
    if (!maxWidth || maxWidth <= 0) return text.split('\n');

    const lines: string[] = [];
    for (const paragraph of String(text).split('\n')) {
        if (!paragraph) {
            lines.push('');
            continue;
        }

        let line = '';
        for (const char of Array.from(paragraph)) {
            const candidate = line + char;
            let width: number;
            try {
                width = font.widthOfTextAtSize(candidate, size);
            } catch {
                // 标准字体遇到缺字形会抛错，降级为「放得下就继续」
                width = candidate.length * size * 0.5;
            }

            if (width <= maxWidth || !line) {
                line = candidate;
                continue;
            }

            // 超宽：在最后一个空格处断行，避免西文单词被劈开
            const lastSpace = line.lastIndexOf(' ');
            if (lastSpace > 0 && /[\x20-\x7E]/.test(char)) {
                lines.push(line.slice(0, lastSpace));
                line = line.slice(lastSpace + 1) + char;
            } else {
                lines.push(line);
                line = char;
            }
        }
        if (line) lines.push(line);
    }
    return lines;
}

/** 图片字节类型识别 */
function detectImageType(bytes: Uint8Array): 'png' | 'jpg' | undefined {
    if (bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'png';
    if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpg';
    return undefined;
}

/** 图片源归一化为字节 */
async function readImageBytes(src: string | ArrayBuffer | Uint8Array): Promise<Uint8Array> {
    return toUint8Array(src as string);
}

/** 绘制上下文 */
interface DrawContext {
    doc: PDFDocument;
    fonts: FontBundle;
    page: ReturnType<PDFDocument['addPage']>;
    pageHeight: number;
    contentWidth: number;
    originX: number;
    originY: number;
    defaultSize: number;
    defaultColor: RGB;
}

/** 在页面上绘制文本（含换行与对齐），返回占用高度 */
function drawTextElement(
    ctx: DrawContext,
    element: PdfTextElement | (Omit<PdfTextElement, 'type'> & { type: 'heading'; level: number })
): number {
    const isHeading = element.type === 'heading';
    const size = element.size ?? (isHeading ? Math.max(10, 24 - ((element.level - 1) as number) * 2.6) : ctx.defaultSize);

    // 标题与加粗走 bold 字体，斜体走 italic，二者同时命中用 boldItalic
    const bold = isHeading || (('bold' in element ? element.bold : false) ?? false);
    const italic = !isHeading && (('italic' in element ? element.italic : false) ?? false);
    const font = bold && italic
        ? ctx.fonts.boldItalic
        : bold
          ? ctx.fonts.bold
          : italic
            ? ctx.fonts.italic
            : ctx.fonts.normal;

    const lineHeight = (element.lineHeight ?? 1.35) * size;
    const width = element.width ?? ctx.contentWidth;
    const lines = wrapText(element.text ?? '', font, size, width);

    let y = ctx.originY + (element.y ?? 0);
    for (const line of lines) {
        let x = ctx.originX + (element.x ?? 0);
        if (element.align === 'center' || element.align === 'right') {
            let lineWidth: number;
            try {
                lineWidth = font.widthOfTextAtSize(line, size);
            } catch {
                lineWidth = line.length * size * 0.5;
            }
            const offset = width - lineWidth;
            x += element.align === 'center' ? offset / 2 : Math.max(0, offset);
        }

        ctx.page.drawText(line, {
            x,
            // 顶部原点 -> pdf-lib 左下原点
            y: ctx.pageHeight - y - size,
            size,
            font,
            color: parseColor(element.color, ctx.defaultColor)
        });
        y += lineHeight;
    }

    return y - (ctx.originY + (element.y ?? 0));
}

/** 绘制表格，返回占用高度 */
function drawTableElement(ctx: DrawContext, element: PdfTableElement): number {
    const fontSize = element.fontSize ?? 10;
    const padding = element.padding ?? 4;
    const hasHeader = element.header !== false && element.rows.length > 0;
    const columnCount = element.rows.reduce((max, row) => Math.max(max, row.length), 0);
    if (!columnCount) return 0;

    const widths = element.columnWidths?.length
        ? element.columnWidths
        : new Array(columnCount).fill(ctx.contentWidth / columnCount);

    const borderColor = element.borderColor === 'none' ? undefined : parseColor(element.borderColor || '#CCCCCC');
    const headerFill = element.headerFill ? parseColor(element.headerFill) : undefined;

    let y = ctx.originY + (element.y ?? 0);
    const startX = ctx.originX + (element.x ?? 0);

    element.rows.forEach((row, rowIndex) => {
        const isHeader = hasHeader && rowIndex === 0;
        // 行高：优先用指定值，否则按最大字号内容估算
        const rowHeight = element.rowHeight ?? fontSize * 1.9;

        if (isHeader && headerFill) {
            ctx.page.drawRectangle({
                x: startX,
                y: ctx.pageHeight - y - rowHeight,
                width: widths.reduce((sum, w) => sum + w, 0),
                height: rowHeight,
                color: headerFill
            });
        }

        let x = startX;
        row.forEach((cell, colIndex) => {
            const cellWidth = widths[colIndex] ?? widths[widths.length - 1] ?? ctx.contentWidth / columnCount;

            if (borderColor) {
                ctx.page.drawRectangle({
                    x,
                    y: ctx.pageHeight - y - rowHeight,
                    width: cellWidth,
                    height: rowHeight,
                    borderColor,
                    borderWidth: 0.5,
                    color: isHeader ? headerFill : undefined
                });
            }

            const text = typeof cell === 'string' ? cell : cellToPlainText(cell);
            if (text) {
                const font = isHeader ? ctx.fonts.bold : ctx.fonts.normal;
                let textWidth: number;
                try {
                    textWidth = font.widthOfTextAtSize(text, fontSize);
                } catch {
                    textWidth = text.length * fontSize * 0.5;
                }
                ctx.page.drawText(text, {
                    x: x + Math.max(padding, (cellWidth - textWidth) / 2),
                    y: ctx.pageHeight - y - rowHeight / 2 - fontSize * 0.35,
                    size: fontSize,
                    font,
                    color: ctx.defaultColor
                });
            }

            x += cellWidth;
        });

        y += rowHeight;
    });

    return y - (ctx.originY + (element.y ?? 0));
}

/** 表格单元格 -> 纯文本（嵌套元素取其文本拼接） */
function cellToPlainText(cell: PdfElement[]): string {
    return cell
        .map((element) => {
            if (element.type === 'text' || element.type === 'heading') return element.text ?? '';
            return '';
        })
        .join(' ')
        .trim();
}

/** 绘制单个元素，返回占用高度 */
async function drawElement(ctx: DrawContext, element: PdfElement): Promise<number> {
    switch (element.type) {
        case 'text':
        case 'heading':
            return drawTextElement(ctx, element as never);

        case 'line': {
            ctx.page.drawLine({
                start: { x: ctx.originX + element.x1, y: ctx.pageHeight - element.y1 },
                end: { x: ctx.originX + element.x2, y: ctx.pageHeight - element.y2 },
                thickness: element.lineWidth ?? 1,
                color: parseColor(element.color, rgb(0.6, 0.6, 0.6)),
                dashArray: element.dash
            });
            return element.lineWidth ?? 1;
        }

        case 'rect': {
            const y = ctx.originY + element.y;
            const options = {
                x: ctx.originX + element.x,
                y: ctx.pageHeight - y - element.height,
                width: element.width,
                height: element.height,
                color: element.fill ? parseColor(element.fill) : undefined,
                borderColor: element.stroke ? parseColor(element.stroke) : undefined,
                borderWidth: element.stroke ? element.lineWidth ?? 1 : 0
            };
            if (element.radius && element.radius > 0) {
                // pdf-lib 没有圆角矩形 API，用 SVG path 绘制
                const { page, pageHeight, originX, originY } = ctx;
                const px = originX + element.x;
                const py = pageHeight - y - element.height;
                const r = Math.min(element.radius, element.width / 2, element.height / 2);
                const path = [
                    `M ${px + r} ${py}`,
                    `L ${px + element.width - r} ${py}`,
                    `A ${r} ${r} 0 0 1 ${px + element.width} ${py + r}`,
                    `L ${px + element.width} ${py + element.height - r}`,
                    `A ${r} ${r} 0 0 1 ${px + element.width - r} ${py + element.height}`,
                    `L ${px + r} ${py + element.height}`,
                    `A ${r} ${r} 0 0 1 ${px} ${py + element.height - r}`,
                    `L ${px} ${py + r}`,
                    `A ${r} ${r} 0 0 1 ${px + r} ${py}`
                ].join(' ');
                page.drawSvgPath(path, {
                    borderColor: options.borderColor,
                    borderWidth: options.borderWidth,
                    color: options.color
                });
            } else {
                ctx.page.drawRectangle(options);
            }
            return element.height;
        }

        case 'image': {
            try {
                const bytes = await readImageBytes(element.src);
                const kind = detectImageType(bytes);
                if (!kind) throw new TypeError('仅支持 PNG / JPEG 图片');
                const image: PDFImage = kind === 'png' ? await ctx.doc.embedPng(bytes) : await ctx.doc.embedJpg(bytes);
                const width = element.width ?? image.width;
                const height = element.height ?? image.height;
                ctx.page.drawImage(image, {
                    x: ctx.originX + (element.x ?? 0),
                    y: ctx.pageHeight - (ctx.originY + (element.y ?? 0)) - height,
                    width,
                    height
                });
                return height;
            } catch {
                return 0;
            }
        }

        case 'table':
            return drawTableElement(ctx, element);

        case 'pageBreak':
            return Number.POSITIVE_INFINITY;

        default:
            return 0;
    }
}

/** 生成 PDF
 * @param spec - 文档规格（页面或元素序列）
 * @param options - 生成选项
 * @returns PDF 字节数据
 */
export async function generatePdf(
    spec: PdfDocumentSpec,
    options: PdfGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        if (!spec || typeof spec !== 'object') throw new TypeError('缺少文档规格');

        const doc = await PDFDocument.create();
        const fonts = await buildFonts(doc, { ...(spec.fonts || {}), ...(options.fonts || {}) });

        // 未内嵌字体时提前拦下中文，避免输出空白方块
        if (!fonts.embedded) {
            for (const element of collectElements(spec)) {
                if ((element.type === 'text' || element.type === 'heading') && hasCjk(element.text ?? '')) {
                    throw new JdocError(
                        '生成中文 PDF 需要内嵌字体：请通过 fonts.regular 传入中文 TTF/OTF 字节',
                        'pdf'
                    );
                }
            }
        }

        const defaultSize = spec.defaultSize ?? 12;
        const defaultColor = parseColor(spec.defaultColor, rgb(0.12, 0.12, 0.12));

        // 元数据
        if (spec.metadata) {
            doc.setTitle(spec.metadata.title || '');
            doc.setAuthor(spec.metadata.author || '');
            doc.setSubject(spec.metadata.subject || '');
            doc.setKeywords(spec.metadata.keywords ? spec.metadata.keywords.split(/[,;\s]+/).filter(Boolean) : []);
            doc.setProducer(spec.metadata.producer || '@fefeding/jdoc');
            doc.setCreator(spec.metadata.creator || '@fefeding/jdoc');
            if (spec.metadata.created) {
                const date = new Date(spec.metadata.created);
                if (!Number.isNaN(date.getTime())) doc.setCreationDate(date);
            }
            if (spec.metadata.modified) {
                const date = new Date(spec.metadata.modified);
                if (!Number.isNaN(date.getTime())) doc.setModificationDate(date);
            }
        }

        /** 显式分页：每页独立坐标系，元素超出页高不再自动拆页 */
        const renderPage = async (elements: PdfElement[], pageSpec: PdfPageSpecLike & { size?: PdfPageSize; orientation?: 'portrait' | 'landscape' }) => {
            const [pageWidth, pageHeight] = resolvePageSize(pageSpec.size, pageSpec.orientation === 'landscape');
            const margin = resolveMargin(pageSpec.margin);
            const page = doc.addPage([pageWidth, pageHeight]);
            const ctx: DrawContext = {
                doc,
                fonts,
                page,
                pageHeight,
                contentWidth: pageWidth - margin.left - margin.right,
                originX: margin.left,
                originY: margin.top,
                defaultSize,
                defaultColor
            };

            let cursorY = 0;
            for (const element of elements) {
                if (element.type === 'pageBreak') {
                    cursorY = Number.POSITIVE_INFINITY;
                    break;
                }
                const used = await drawElement({ ...ctx, originY: margin.top + cursorY }, element);
                cursorY += used;
            }
        };

        if (spec.pages?.length) {
            for (const pageSpec of spec.pages) {
                await renderPage(pageSpec.elements || spec.elements || [], pageSpec);
            }
        } else {
            // 自动分页：内容超出可用高度时另起一页
            const elements = spec.elements || [];
            const [pageWidth, pageHeight] = resolvePageSize(undefined, false);
            const margin = resolveMargin(undefined);
            const usableHeight = pageHeight - margin.top - margin.bottom;

            let page = doc.addPage([pageWidth, pageHeight]);
            let ctx: DrawContext = {
                doc,
                fonts,
                page,
                pageHeight,
                contentWidth: pageWidth - margin.left - margin.right,
                originX: margin.left,
                originY: margin.top,
                defaultSize,
                defaultColor
            };
            let cursorY = 0;

            for (const element of elements) {
                if (element.type === 'pageBreak') {
                    page = doc.addPage([pageWidth, pageHeight]);
                    ctx = { ...ctx, page, originY: margin.top };
                    cursorY = 0;
                    continue;
                }

                const used = await drawElement({ ...ctx, originY: margin.top + cursorY }, element);
                if (cursorY + used > usableHeight && cursorY > 0) {
                    // 放不下就换页重绘一次（只重试一次，避免元素本身高于整页时死循环）
                    page = doc.addPage([pageWidth, pageHeight]);
                    ctx = { ...ctx, page, originY: margin.top };
                    cursorY = await drawElement({ ...ctx, originY: margin.top }, element);
                } else {
                    cursorY += used;
                }
            }
        }

        const bytes = await doc.save({ useObjectStreams: options.compress !== false });
        return bytes;
    } catch (error) {
        throw wrapError('pdf', '生成', error);
    }
}

/** 收集文档中所有元素（供前置校验使用） */
function collectElements(spec: PdfDocumentSpec): PdfElement[] {
    if (spec.pages?.length) return spec.pages.flatMap((page) => page.elements || []);
    return spec.elements || [];
}
