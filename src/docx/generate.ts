/**
 * DOCX 生成
 *
 * 基于 docx 库把 jdoc 的块结构写回 OOXML。设计原则：JSON 结构即 API，
 * 所有必填项都给字符串简写（text / runs 二选一），减少调用方样板代码。
 *
 * @module docx/generate
 */

import {
    AlignmentType,
    BorderStyle,
    Document,
    ExternalHyperlink,
    HeadingLevel,
    ImageRun,
    LevelFormat,
    PageBreak,
    Packer,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    WidthType,
    type IImageOptions
} from 'docx';
import { JdocError } from '../core/errors';
import { bytesToBase64, toUint8Array, type DocInput } from '../core/input';
import type {
    DocxBlock,
    DocxCell,
    DocxDocument,
    DocxGenerateOptions,
    DocxList,
    DocxRun,
    DocxTable
} from './types';

/** 磅转 dxa（1 磅 = 20 dxa） */
const DXA_PER_PT = 20;

/** 磅转半磅（docx 库字号单位） */
const HALF_PT_PER_PT = 2;

/** 纸张尺寸（dxa） */
const PAGE_SIZES: Record<string, { width: number; height: number }> = {
    A3: { width: 16838, height: 23811 },
    A4: { width: 11906, height: 16838 },
    A5: { width: 8391, height: 11906 },
    Letter: { width: 12240, height: 15840 },
    Legal: { width: 12240, height: 20160 }
};

/** 引用块段落样式 ID（样式名固定为 Quote，供解析端识别） */
const QUOTE_STYLE_ID = 'JdocQuote';

/** 有序列表 numbering reference */
const ORDERED_REF = 'jdoc-ordered';
/** 无序列表 numbering reference */
const UNORDERED_REF = 'jdoc-unordered';

/** 列表各层级缩进（dxa） */
const LIST_INDENT = [
    { left: 720, hanging: 360 },
    { left: 1440, hanging: 360 },
    { left: 2160, hanging: 360 },
    { left: 2880, hanging: 360 }
];

/** 颜色值规范化：去掉 # 前缀，docx 库要求裸 hex */
function bareColor(value: string | undefined): string | undefined {
    if (!value) return undefined;
    const hex = value.replace(/^#/, '').trim();
    return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : undefined;
}

/** DocxRun -> TextRun（含超链接拆分为 ExternalHyperlink） */
function buildTextRun(run: DocxRun, style: { font?: string; size?: number }): TextRun | ExternalHyperlink {
    const options: Record<string, unknown> = {
        text: run.text,
        bold: run.bold,
        italics: run.italic,
        underline: run.underline ? {} : undefined,
        strike: run.strike,
        subScript: run.subscript,
        superScript: run.superscript,
        color: bareColor(run.color),
        // docx 库字号单位是半磅
        size: run.size !== undefined ? Math.round(run.size * HALF_PT_PER_PT) : style.size !== undefined ? Math.round(style.size * HALF_PT_PER_PT) : undefined,
        font: run.font || style.font,
        highlight: run.highlight ? bareColor(run.highlight) : undefined
    };

    const textRun = new TextRun(options as TextRunOptionsLike);

    if (run.link) return new ExternalHyperlink({ children: [textRun], link: run.link });
    return textRun;
}

/** TextRun 构造参数（避免直接依赖 docx 的类型导出细节） */
type TextRunOptionsLike = ConstructorParameters<typeof TextRun>[0];

/** 从 DocxRun[] 或 text 字符串构造段落子节点 */
function buildChildren(
    text: string | undefined,
    runs: DocxRun[] | undefined,
    style: { font?: string; size?: number }
): (TextRun | ExternalHyperlink)[] {
    if (runs && runs.length) return runs.map((run) => buildTextRun(run, style));
    if (text !== undefined && text !== '') return [new TextRun({ text, font: style.font, size: style.size !== undefined ? style.size * HALF_PT_PER_PT : undefined } as TextRunOptionsLike)];
    return [];
}

/** 缩进与对齐 -> 段落配置 */
function buildParagraphLayout(block: DocxBlock): Record<string, unknown> {
    const layout: Record<string, unknown> = {};
    const para = block as { align?: string; indentLeft?: number; indentRight?: number; indentFirstLine?: number; spacingBefore?: number; spacingAfter?: number; lineSpacing?: number };

    if (para.align) layout.alignment = AlignmentType[para.align.toUpperCase() as keyof typeof AlignmentType];

    const indent: Record<string, number> = {};
    if (para.indentLeft) indent.left = Math.round(para.indentLeft * DXA_PER_PT);
    if (para.indentRight) indent.right = Math.round(para.indentRight * DXA_PER_PT);
    // 悬挂缩进在 OOXML 中以 hanging 的正值表示
    if (para.indentFirstLine !== undefined) {
        if (para.indentFirstLine < 0) indent.hanging = Math.round(-para.indentFirstLine * DXA_PER_PT);
        else indent.firstLine = Math.round(para.indentFirstLine * DXA_PER_PT);
    }
    if (Object.keys(indent).length) layout.indent = indent;

    const spacing: Record<string, number> = {};
    if (para.spacingBefore) spacing.before = Math.round(para.spacingBefore * DXA_PER_PT);
    if (para.spacingAfter) spacing.after = Math.round(para.spacingAfter * DXA_PER_PT);
    // 行距倍数以 240 = 单倍行距表达
    if (para.lineSpacing && para.lineSpacing !== 1) spacing.line = Math.round(para.lineSpacing * 240);
    if (Object.keys(spacing).length) layout.spacing = spacing;

    return layout;
}

/** data URL / base64 字符串 -> { bytes, mime } */
async function decodeImageSource(
    src: string | ArrayBuffer | Uint8Array
): Promise<{ bytes: Uint8Array; mime: string }> {
    if (typeof src !== 'string') {
        return { bytes: await toUint8Array(src), mime: 'png' };
    }

    const match = /^data:([^;,]+)(;base64)?,([\s\S]*)$/.exec(src.trim());
    if (match) {
        const mime = match[1] || 'image/png';
        const payload = match[3] || '';
        const bytes = match[2] ? decodeBase64(payload) : new TextEncoder().encode(decodeURIComponent(payload));
        return { bytes, mime };
    }

    return { bytes: decodeBase64(src), mime: 'png' };
}

/** base64 -> Uint8Array（双端） */
function decodeBase64(value: string): Uint8Array {
    const clean = value.replace(/\s+/g, '');
    const atobFn = (globalThis as { atob?: (s: string) => string }).atob;
    const binary = atobFn ? atobFn(clean) : Buffer.from(clean, 'base64').toString('binary');
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i) & 0xff;
    return bytes;
}

/** MIME -> docx 库要求的图片类型 */
function toImageType(mime: string): NonNullable<IImageOptions['type']> {
    const m = mime.toLowerCase();
    if (m.includes('jpeg') || m.includes('jpg')) return 'jpg';
    if (m.includes('gif')) return 'gif';
    if (m.includes('bmp')) return 'bmp';
    if (m.includes('svg')) return 'svg';
    return 'png';
}

/** 磅 -> 像素（docx 的 ImageRun 尺寸单位是 px，按 96dpi 折算） */
function ptToPx(pt: number | undefined, fallback: number): number {
    if (!pt || pt <= 0) return fallback;
    return Math.round((pt * 96) / 72);
}

/** 图片块 -> 段落数组（图片 + 可选图注） */
async function buildImageParagraphs(
    block: Extract<DocxBlock, { type: 'image' }>,
    style: { font?: string; size?: number }
): Promise<Paragraph[]> {
    const paragraphs: Paragraph[] = [];
    try {
        const { bytes, mime } = await decodeImageSource(block.src);
        // 未指定尺寸时用 300x150 的常见默认，避免 docx 库因 0 尺寸报错
        const width = ptToPx(block.width, 300);
        const height = ptToPx(block.height, Math.round((300 / 4) * 3));

        const imageRun = new ImageRun({
            type: toImageType(mime),
            // 直接传 Uint8Array，避免 Buffer（浏览器/非 Node 环境未定义）
            data: bytes,
            transformation: { width, height },
            altText: block.alt ? { name: block.alt, title: block.alt, description: block.alt } : undefined
        } as ConstructorParameters<typeof ImageRun>[0]);

        paragraphs.push(new Paragraph({ children: [imageRun], alignment: AlignmentType.CENTER }));
    } catch {
        // 图片损坏不阻断文档生成，降级为占位段落
        paragraphs.push(new Paragraph({ children: [new TextRun({ text: '[图片无法解析]', italics: true, color: '999999' } as TextRunOptionsLike)] }));
    }

    if (block.caption) {
        paragraphs.push(
            new Paragraph({
                children: [new TextRun({ text: block.caption, size: 18, color: '666666' } as TextRunOptionsLike)],
                alignment: AlignmentType.CENTER
            })
        );
    }

    return paragraphs;
}

/** 列表块 -> 段落数组（递归处理子项层级） */
function buildListParagraphs(block: DocxList, level = 0): Paragraph[] {
    const paragraphs: Paragraph[] = [];
    const indent = LIST_INDENT[Math.min(level, LIST_INDENT.length - 1)];

    for (const item of block.items) {
        const children = typeof item === 'string' ? undefined : (item as DocxRun[]);
        const childItems = typeof item === 'string' ? [] : ((item as { children?: unknown[] }).children || []);

        paragraphs.push(
            new Paragraph({
                children: buildChildren(typeof item === 'string' ? item : undefined, children, {}),
                numbering: {
                    reference: block.ordered ? ORDERED_REF : UNORDERED_REF,
                    level: Math.min(level, 8)
                },
                indent
            })
        );

        // 子项缩进一级继续输出
        if (childItems.length) {
            paragraphs.push(
                ...buildListParagraphs({ type: 'list', ordered: block.ordered, items: childItems as never[] }, level + 1)
            );
        }
    }

    return paragraphs;
}

/** 表格单元格 -> TableCell */
function buildTableCell(
    cell: DocxCell,
    isHeader: boolean,
    borderSize: number
): TableCell {
    // 字符串单元格需要显式包成 run；此前非表头字符串会走到空 run 分支，导致内容丢失
    const base: DocxRun[] =
        typeof cell === 'string' ? (cell ? [{ text: cell }] : []) : (cell || []).filter(Boolean);
    const styled = isHeader ? base.map((run) => ({ ...run, bold: true })) : base;

    return new TableCell({
        children: [
            new Paragraph({
                children: styled.length
                    ? styled.map((run) => buildTextRun(run, {}))
                    : [new TextRun({ text: '' } as TextRunOptionsLike)]
            })
        ],
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        borders: {
            top: { style: borderSize > 0 ? BorderStyle.SINGLE : BorderStyle.NONE, size: Math.max(2, Math.round(borderSize * 8)), color: 'CCCCCC' },
            bottom: { style: borderSize > 0 ? BorderStyle.SINGLE : BorderStyle.NONE, size: Math.max(2, Math.round(borderSize * 8)), color: 'CCCCCC' },
            left: { style: borderSize > 0 ? BorderStyle.SINGLE : BorderStyle.NONE, size: Math.max(2, Math.round(borderSize * 8)), color: 'CCCCCC' },
            right: { style: borderSize > 0 ? BorderStyle.SINGLE : BorderStyle.NONE, size: Math.max(2, Math.round(borderSize * 8)), color: 'CCCCCC' }
        }
    });
}

/** 表格块 -> Table */
function buildTable(block: DocxTable): Table {
    const borderSize = block.borderSize === undefined ? 1 : block.borderSize;
    const hasHeader = block.header !== false && block.rows.length > 0;
    const columnCount = block.rows.reduce((max, row) => Math.max(max, row.length), 0);

    const rows = block.rows.map(
        (row, rowIndex) =>
            new TableRow({
                // 首行作为表头时设置重复标题，打印多页也能带上表头
                tableHeader: hasHeader && rowIndex === 0,
                children: Array.from({ length: columnCount }, (_, colIndex) =>
                    buildTableCell(row[colIndex] ?? '', hasHeader && rowIndex === 0, borderSize)
                )
            })
    );

    return new Table({
        rows,
        width: block.columnWidths
            ? { size: Math.round(block.columnWidths.reduce((sum, w) => sum + w, 0) * DXA_PER_PT), type: WidthType.DXA }
            : { size: 100, type: WidthType.PERCENTAGE }
    });
}

/** 标题层级 -> docx HeadingLevel */
const HEADING_MAP: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6
};

/** 文档 body 的子节点类型：段落或表格 */
type DocxChild = Paragraph | Table;

/** 单个块 -> 子节点数组（部分块会展开为多段，如图片+图注） */
async function buildBlockParagraphs(
    block: DocxBlock,
    style: { font?: string; size?: number }
): Promise<DocxChild[]> {
    switch (block.type) {
        case 'heading': {
            const layout = buildParagraphLayout(block);
            const heading = HEADING_MAP[block.level] || HeadingLevel.HEADING_1;
            const children = buildChildren(block.text, block.runs, style);
            // 标题未显式指定时沿用调用方样式，允许用 bold 模拟一级标题
            return [new Paragraph({ children, heading, ...layout })];
        }

        case 'paragraph': {
            const layout = buildParagraphLayout(block);
            return [new Paragraph({ children: buildChildren(block.text, block.runs, style), ...layout })];
        }

        case 'quote': {
            // 在构造 run 之前就带上斜体；把已构造的 TextRun 实例展开重建会丢掉 text 等属性
            const base: DocxRun[] =
                block.runs && block.runs.length
                    ? block.runs
                    : block.text !== undefined
                      ? [{ text: block.text }]
                      : [];
            return [
                new Paragraph({
                    // 套用名为 Quote 的段落样式，解析端据此还原成 quote 块（仅靠边框无法识别）
                    style: QUOTE_STYLE_ID,
                    children: base.map((run) => buildTextRun({ ...run, italic: true }, style)),
                    indent: { left: 360 },
                    border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'CCCCCC', space: 8 } }
                })
            ];
        }

        case 'code': {
            // 代码块按原始换行拆成多段，保留缩进与等宽字体
            const lines = String(block.text ?? '').split('\n');
            const font = 'Consolas';
            return lines.map(
                (line) =>
                    new Paragraph({
                        children: [new TextRun({ text: line || ' ', font, size: 20 } as TextRunOptionsLike)],
                        shading: { type: 'clear', fill: 'F5F5F5', color: 'auto' },
                        spacing: { before: 0, after: 0 }
                    })
            );
        }

        case 'list':
            return buildListParagraphs(block);

        case 'table':
            return [buildTable(block)];

        case 'image':
            return buildImageParagraphs(block, style);

        case 'pageBreak':
            return [new Paragraph({ children: [new PageBreak()] })];

        case 'hr':
            return [
                new Paragraph({
                    children: [],
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'DDDDDD', space: 1 } },
                    spacing: { before: 120, after: 120 }
                })
            ];

        default:
            return [];
    }
}

/** docx 库接受的编号格式枚举值 */
type LevelFormatValue = (typeof LevelFormat)[keyof typeof LevelFormat];

/** 生成 docx 所需的 numbering 配置（有序 + 无序各 4 级） */
function buildNumberingConfig() {
    const levels = (format: LevelFormatValue, text: string) =>
        LIST_INDENT.map((indent, level) => ({
            level,
            format,
            text: format === LevelFormat.BULLET ? text : `%${level + 1}.`,
            alignment: AlignmentType.START,
            style: { paragraph: { indent } }
        }));

    return [
        { reference: ORDERED_REF, levels: levels(LevelFormat.DECIMAL, '%1.') },
        { reference: UNORDERED_REF, levels: levels(LevelFormat.BULLET, '•') }
    ];
}

/** 归一化页边距 */
function buildMargin(margin: DocxGenerateOptions['margin']) {
    const DEFAULT = 1440; // 1 英寸
    if (typeof margin === 'number') {
        const value = Math.round(margin * DXA_PER_PT);
        return { top: value, right: value, bottom: value, left: value };
    }
    if (margin && typeof margin === 'object') {
        return {
            top: Math.round((margin.top ?? 72) * DXA_PER_PT),
            right: Math.round((margin.right ?? 72) * DXA_PER_PT),
            bottom: Math.round((margin.bottom ?? 72) * DXA_PER_PT),
            left: Math.round((margin.left ?? 72) * DXA_PER_PT)
        };
    }
    return { top: DEFAULT, right: DEFAULT, bottom: DEFAULT, left: DEFAULT };
}

/**
 * 生成 DOCX 文件
 * @param input - DocxDocument，或直接传 blocks 数组（简写）
 * @param options - 页面与默认样式
 * @returns DOCX 字节数据
 */
export async function generateDocx(
    input: DocxDocument | DocxBlock[],
    options: DocxGenerateOptions = {}
): Promise<Uint8Array> {
    const doc: DocxDocument = Array.isArray(input) ? { blocks: input } : input;
    if (!doc || !Array.isArray(doc.blocks)) {
        throw new JdocError('生成 DOCX 失败：缺少 blocks 内容', 'docx');
    }

    const style = {
        font: options.defaultFont,
        size: options.defaultSize !== undefined ? Math.round(options.defaultSize * HALF_PT_PER_PT) : undefined
    };

    const children: DocxChild[] = [];
    for (const block of doc.blocks) {
        if (!block || typeof block !== 'object') continue;
        children.push(...(await buildBlockParagraphs(block, style)));
    }

    // 空文档也需要至少一个段落，否则 Word 会判定文件损坏
    if (!children.length) children.push(new Paragraph({ children: [] }));

    const landscape = options.orientation === 'landscape';
    const sizeKey = typeof options.pageSize === 'string' ? options.pageSize : 'A4';
    const rawSize =
        typeof options.pageSize === 'object' && options.pageSize
            ? {
                  width: Math.round(options.pageSize.width * DXA_PER_PT),
                  height: Math.round(options.pageSize.height * DXA_PER_PT)
              }
            : PAGE_SIZES[sizeKey] || PAGE_SIZES.A4;
    const size = landscape
        ? { width: rawSize.height, height: rawSize.width, orientation: 'landscape' as const }
        : { width: rawSize.width, height: rawSize.height, orientation: 'portrait' as const };

    const document = new Document({
        // 声明引用样式：解析端按样式名 Quote 还原 quote 块
        styles: {
            paragraphStyles: [
                {
                    id: QUOTE_STYLE_ID,
                    name: 'Quote',
                    basedOn: 'Normal',
                    quickFormat: true,
                    run: { italics: true, color: '4A4A4A' },
                    paragraph: { indent: { left: 360 }, spacing: { before: 120, after: 120 } }
                }
            ]
        },
        creator: doc.metadata?.creator || doc.metadata?.author || 'jdoc',
        title: doc.metadata?.title,
        subject: doc.metadata?.subject,
        description: doc.metadata?.description,
        keywords: doc.metadata?.keywords,
        lastModifiedBy: doc.metadata?.lastModifiedBy,
        numbering: { config: buildNumberingConfig() },
        sections: [
            {
                properties: {
                    page: {
                        size,
                        margin: buildMargin(options.margin)
                    }
                },
                children
            }
        ]
    });

    // 使用 toArrayBuffer 以跨平台：toBuffer 在非 Node 环境（浏览器、Deno、部分
    // 打包运行时）会抛 "nodebuffer is not supported by this platform"。
    const arrayBuffer = await Packer.toArrayBuffer(document);
    return new Uint8Array(arrayBuffer);
}
