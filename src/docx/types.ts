/**
 * DOCX 文档结构类型定义
 *
 * 解析（docxToJson）与生成（jsonToDocx）共用这一套结构，
 * 使 DOCX 也能像 PPTX 一样做「解析 → JSON → 生成」的往返。
 *
 * @module docx/types
 */

/** 文本级样式（对应一次 run） */
export interface DocxRun {
    /** 文本内容 */
    text: string;
    /** 加粗 */
    bold?: boolean;
    /** 斜体 */
    italic?: boolean;
    /** 下划线 */
    underline?: boolean;
    /** 删除线 */
    strike?: boolean;
    /** 下标 */
    subscript?: boolean;
    /** 上标 */
    superscript?: boolean;
    /** 文字颜色，#RRGGBB */
    color?: string;
    /** 字号（磅），如 12 表示 12pt */
    size?: number;
    /** 字体名 */
    font?: string;
    /** 背景高亮色，#RRGGBB */
    highlight?: string;
    /** 超链接地址 */
    link?: string;
}

/** 段落对齐方式 */
export type DocxAlign = 'left' | 'center' | 'right' | 'justify';

/** 标题块 */
export interface DocxHeading {
    type: 'heading';
    /** 标题层级 1-6 */
    level: 1 | 2 | 3 | 4 | 5 | 6;
    /** 纯文本简写 */
    text?: string;
    /** 富文本（与 text 二选一，text 优先） */
    runs?: DocxRun[];
}

/** 普通段落块 */
export interface DocxParagraph {
    type: 'paragraph';
    text?: string;
    runs?: DocxRun[];
    align?: DocxAlign;
    /** 段落缩进（磅） */
    indentLeft?: number;
    indentRight?: number;
    /** 首行缩进（磅） */
    indentFirstLine?: number;
    /** 段前间距（磅） */
    spacingBefore?: number;
    /** 段后间距（磅） */
    spacingAfter?: number;
    /** 行距倍数，1 为单倍行距 */
    lineSpacing?: number;
}

/** 列表项：字符串简写或带子项的富文本 */
export type DocxListItem = string | (DocxRun[] & { children?: DocxListItem[] });

/** 列表块 */
export interface DocxList {
    type: 'list';
    /** true 为有序（编号）列表，false 为无序（项目符号）列表 */
    ordered?: boolean;
    /** 列表项 */
    items: DocxListItem[];
}

/** 表格单元格：字符串简写或富文本 */
export type DocxCell = string | DocxRun[];

/** 表格块 */
export interface DocxTable {
    type: 'table';
    /** 二维单元格内容，行优先 */
    rows: DocxCell[][];
    /** 首行是否为表头（生成时加粗并重复表头） */
    header?: boolean;
    /** 列宽（磅），长度需与列数一致 */
    columnWidths?: number[];
    /** 表格边框宽度（磅），0 表示无边框 */
    borderSize?: number;
}

/** 图片块 */
export interface DocxImage {
    type: 'image';
    /** 图片数据：data URL / base64 / 二进制 */
    src: string | ArrayBuffer | Uint8Array;
    /** 宽度（磅） */
    width?: number;
    /** 高度（磅） */
    height?: number;
    /** 图注文字 */
    caption?: string;
    /** 图片说明（alt 文本） */
    alt?: string;
}

/** 分页符 */
export interface DocxPageBreak {
    type: 'pageBreak';
}

/** 分隔线（水平线） */
export interface DocxHorizontalRule {
    type: 'hr';
}

/** 引用块 */
export interface DocxQuote {
    type: 'quote';
    text?: string;
    runs?: DocxRun[];
}

/** 代码块（等宽字体 + 底纹） */
export interface DocxCode {
    type: 'code';
    text: string;
    /** 代码语言名（仅作标注） */
    language?: string;
}

/** 所有支持的块类型 */
export type DocxBlock =
    | DocxHeading
    | DocxParagraph
    | DocxList
    | DocxTable
    | DocxImage
    | DocxPageBreak
    | DocxHorizontalRule
    | DocxQuote
    | DocxCode;

/** 文档元数据 */
export interface DocxMetadata {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    creator?: string;
    lastModifiedBy?: string;
    company?: string;
}

/** 完整 DOCX 文档（JSON 契约） */
export interface DocxDocument {
    /** 文档元数据 */
    metadata?: DocxMetadata;
    /** 内容块序列 */
    blocks: DocxBlock[];
}

/** 解析选项 */
export interface DocxParseOptions {
    /**
     * 自定义 mammoth styleMap，把 Word 内置样式映射到语义标签。
     * 传入的映射会与内置映射合并。
     */
    styleMap?: string[] | Record<string, string>;
    /** 是否使用 mammoth 内置样式映射（默认 true） */
    includeDefaultStyleMap?: boolean;
    /** 图片转换策略：'data-url'（内联 base64）或 'ignore'（丢弃） */
    convertImage?: 'data-url' | 'ignore';
}

/** 解析结果（含警告信息） */
export interface DocxParseResult<T> {
    /** 解析产物 */
    value: T;
    /** mammoth 转换过程中的警告 */
    messages: string[];
    /** 纯文本内容（便于全文检索） */
    rawText?: string;
}

/** 生成选项 */
export interface DocxGenerateOptions {
    /** 纸张尺寸，默认 A4 */
    pageSize?: 'A4' | 'Letter' | 'A3' | 'A5' | 'Legal' | { width: number; height: number };
    /** 页面方向，默认纵向 */
    orientation?: 'portrait' | 'landscape';
    /** 页边距（磅），默认 72（1 英寸） */
    margin?: number | { top?: number; right?: number; bottom?: number; left?: number };
    /** 默认正文字体 */
    defaultFont?: string;
    /** 默认字号（磅） */
    defaultSize?: number;
}