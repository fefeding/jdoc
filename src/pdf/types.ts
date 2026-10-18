import type { AssetSource, DocAsset, DocAssetMode } from '../core/asset';

/**
 * PDF 文档结构类型定义
 *
 * 解析输出为 PdfDocument（含逐页文本与坐标），生成输入为 PdfDocumentSpec（含元素序列）。
 *
 * @module pdf/types
 */

/** 文档元数据 */
export interface PdfMetadata {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    creator?: string;
    producer?: string;
    /** 创建时间（ISO 字符串） */
    created?: string;
    /** 修改时间（ISO 字符串） */
    modified?: string;
    /** PDF 版本，如 '1.7' */
    version?: string;
}

/** 文本片段（带坐标，PDF 坐标系：原点左下、单位 pt） */
export interface PdfTextItem {
    /** 文本内容 */
    text: string;
    /** 左下角 X */
    x: number;
    /** 左下角 Y */
    y: number;
    /** 宽度 */
    width: number;
    /** 高度 */
    height: number;
    /** 字号 */
    fontSize: number;
    /** 字体名 */
    fontName?: string;
    /** 该片段后是否紧跟换行 */
    hasEOL?: boolean;
}

/** 页面 */
export interface PdfPage {
    /** 页号（1 起） */
    index: number;
    /** 宽度（pt） */
    width: number;
    /** 高度（pt） */
    height: number;
    /** 旋转角度（0/90/180/270） */
    rotation: number;
    /** 整页纯文本 */
    text?: string;
    /** 文本片段（含坐标） */
    items?: PdfTextItem[];
}

/** 解析结果文档 */
export interface PdfDocument {
    metadata?: PdfMetadata;
    /** 总页数 */
    pageCount: number;
    /** 页面详情（仅当请求 detailed 时返回） */
    pages?: PdfPage[];
    /**
     * 内嵌资源（附件 / 页面图像），仅当 `assets` 选项非 `ignore` 时返回
     *
     * 附件锚点为 `kind: 'other'`，页面图像为 `kind: 'page'` 并记录页码。
     */
    assets?: DocAsset[];
}

/** 解析选项 */
export interface PdfParseOptions {
    /**
     * 内嵌资源提取方式，默认 `ignore`（不提取）
     *
     * 页面图像需要走渲染指令，代价较高（上百页文档会明显变慢），
     * 因此只有显式开启才提取。附件的提取很廉价，开了就一起给。
     */
    assets?: DocAssetMode;
    /** 是否返回逐页文本，默认 true */
    text?: boolean;
    /** 是否返回带坐标的文本片段，默认 false（体积较大） */
    items?: boolean;
    /** 只解析指定页（1 起），缺省全部 */
    pages?: number[];
    /** 最大返回页数（配合分页预览场景） */
    maxPages?: number;
    /** 中文字体资源目录 URL/路径；缺省自动探测 pdfjs-dist 内的标准字体 */
    standardFontDataUrl?: string;
}

/** 解析结果 */
export interface PdfResult {
    /** 文档 */
    document: PdfDocument;
    /** 解析提示 */
    messages: string[];
}

/** 纸张尺寸（pt） */
export type PdfPageSize = 'A4' | 'A3' | 'A5' | 'Letter' | 'Legal' | [number, number] | { width: number; height: number };

/** 文本对齐 */
export type PdfAlign = 'left' | 'center' | 'right';

/** 文本元素 */
export interface PdfTextElement {
    type: 'text';
    text: string;
    /** 距左边距（pt，默认 0） */
    x?: number;
    /** 距顶边距（pt，默认 0） */
    y?: number;
    /** 字号（pt，默认 12） */
    size?: number;
    /** 字体名（需已注册，默认随 bold/italic 变化） */
    font?: string;
    bold?: boolean;
    italic?: boolean;
    /** 颜色，#RRGGBB 或 'black'/'red' 等关键字 */
    color?: string;
    align?: PdfAlign;
    /** 文本框宽度（pt），给定后支持换行与对齐 */
    width?: number;
    /** 行高倍数，默认 1.35 */
    lineHeight?: number;
}

/** 标题元素 */
export interface PdfHeadingElement extends Omit<PdfTextElement, 'type'> {
    type: 'heading';
    /** 标题层级 1-6 */
    level: 1 | 2 | 3 | 4 | 5 | 6;
}

/** 图片元素 */
export interface PdfImageElement {
    type: 'image';
    /**
     * 图片数据：data URL / base64 / 二进制，或资源引用 `{ assetId }`
     *
     * 引用写法需配合 `jsonToPdf(doc, { assets })`，
     * 适合「解析 → 改 → 生成」时直接复用解析出的字节，不必转成 data URL。
     */
    src: string | ArrayBuffer | Uint8Array | AssetSource;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
}

/** 线段元素 */
export interface PdfLineElement {
    type: 'line';
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    /** 线宽（pt），默认 1 */
    lineWidth?: number;
    color?: string;
    /** 虚线间隔，如 [3, 2] */
    dash?: number[];
}

/** 矩形元素 */
export interface PdfRectElement {
    type: 'rect';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 填充色 */
    fill?: string;
    /** 描边色 */
    stroke?: string;
    /** 描边宽度（pt） */
    lineWidth?: number;
    /** 圆角半径（pt） */
    radius?: number;
}

/** 表格单元格 */
export type PdfTableCell = string | PdfElement[];

/** 表格元素 */
export interface PdfTableElement {
    type: 'table';
    x?: number;
    y?: number;
    /** 各列宽度（pt），缺省按内容区等分 */
    columnWidths?: number[];
    /** 行数据（行优先） */
    rows: PdfTableCell[][];
    /** 首行是否作为表头（默认 true） */
    header?: boolean;
    /** 单元格字号（pt），默认 10 */
    fontSize?: number;
    /** 单元格内边距（pt），默认 4 */
    padding?: number;
    /** 行高（pt），缺省按字号自适应 */
    rowHeight?: number;
    /** 边框色，设为 'none' 则无边框 */
    borderColor?: string;
    /** 表头背景色 */
    headerFill?: string;
}

/** 分页符 */
export interface PdfPageBreakElement {
    type: 'pageBreak';
}

/** PDF 元素联合类型 */
export type PdfElement =
    | PdfTextElement
    | PdfHeadingElement
    | PdfImageElement
    | PdfLineElement
    | PdfRectElement
    | PdfTableElement
    | PdfPageBreakElement;

/** 页面定义 */
export interface PdfPageSpec {
    /** 纸张尺寸，默认 A4 */
    size?: PdfPageSize;
    /** 方向，默认纵向 */
    orientation?: 'portrait' | 'landscape';
    /** 页边距（pt），默认 56（约 2cm） */
    margin?: number | { top?: number; right?: number; bottom?: number; left?: number };
    /** 本页元素；缺省时使用文档级 elements */
    elements?: PdfElement[];
}

/** 生成文档规格 */
export interface PdfDocumentSpec {
    /**
     * 可用资源表：用于解析图片元素的 `{ assetId }` 引用
     *
     * 与 `jsonToPdf` 的 `assets` 选项同效，规格内自带时无需再传。
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
    metadata?: PdfMetadata;
    /** 页面；缺省时按 elements 自动分页 */
    pages?: PdfPageSpec[];
    /** 文档级元素（自动分页时使用） */
    elements?: PdfElement[];
    /** 内嵌字体：{ regular, bold, italic, boldItalic }，用于中文或自定义字体 */
    fonts?: PdfFontSet;
    /** 默认字号（pt），默认 12 */
    defaultSize?: number;
    /** 默认正文颜色 */
    defaultColor?: string;
}

/** 内嵌字体集合 */
export interface PdfFontSet {
    regular?: string | ArrayBuffer | Uint8Array;
    bold?: string | ArrayBuffer | Uint8Array;
    italic?: string | ArrayBuffer | Uint8Array;
    boldItalic?: string | ArrayBuffer | Uint8Array;
}

/** 生成选项 */
export interface PdfGenerateOptions {
    /**
     * 可用资源表：用于解析图片元素上的 `{ assetId }` 引用
     *
     * @example
     * ```js
     * const parsed = await pdfToJson(file, { assets: 'extract' });
     * const bytes  = await jsonToPdf(parsed, { assets: parsed.assets });
     * ```
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
    /** 是否压缩内容流，默认 true */
    compress?: boolean;
    /** 中文/自定义字体，与文档级 fonts 合并 */
    fonts?: PdfFontSet;
}
