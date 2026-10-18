import type { AssetSource, DocAsset, DocAssetMode } from '../core/asset';

/**
 * XLSX 工作簿结构类型定义
 *
 * 解析（xlsxToJson）与生成（jsonToXlsx）共用。为降低使用门槛，
 * 单元格既支持「直接写值」的简写形式，也支持带样式的对象形式。
 *
 * @module xlsx/types
 */

/** 单元格可承载的值类型 */
export type XlsxCellValue = string | number | boolean | Date | null;

/** 富文本片段 */
export interface XlsxRichText {
    text: string;
    font?: XlsxFontStyle;
}

/** 字体样式 */
export interface XlsxFontStyle {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    /** 字体颜色 #RRGGBB */
    color?: string;
}

/** 边框样式 */
export interface XlsxBorderStyle {
    style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'none';
    /** 边框颜色 #RRGGBB */
    color?: string;
}

/** 单元格样式 */
export interface XlsxCellStyle {
    font?: XlsxFontStyle;
    /** 背景填充色 #RRGGBB */
    fill?: string;
    align?: 'left' | 'center' | 'right';
    valign?: 'top' | 'middle' | 'bottom';
    /** 自动换行 */
    wrapText?: boolean;
    /** 数字格式，如 '0.00%'、'yyyy-mm-dd' */
    numFmt?: string;
    border?: {
        top?: XlsxBorderStyle;
        right?: XlsxBorderStyle;
        bottom?: XlsxBorderStyle;
        left?: XlsxBorderStyle;
    };
}

/** 单元格对象形式 */
export interface XlsxCellObject {
    /** 单元格值；与 formula 二选一 */
    value?: XlsxCellValue;
    /** 富文本片段数组 */
    richText?: XlsxRichText[];
    /** 公式（不含前导 =） */
    formula?: string;
    /** 公式缓存结果 */
    result?: XlsxCellValue;
    /** 显示文本简写 */
    text?: string;
    /** 超链接 */
    hyperlink?: string;
    /** 单元格样式 */
    style?: XlsxCellStyle;
    /** 注释 */
    note?: string;
    /**
     * 单元格内嵌图片
     *
     * 可给 data URL / base64 / 二进制，或资源引用 `{ assetId }`（配合工作簿的 `assets`）。
     * 图片会按单元格锚定，与 Excel 中的浮动图片一致。
     */
    image?: AssetSource;
    /** 图片资源 ID（存放在工作簿 `assets` 中时给出） */
    assetId?: string;
    /** 图片显示宽度（磅） */
    imageWidth?: number;
    /** 图片显示高度（磅） */
    imageHeight?: number;
}

/** 单元格：值简写或对象形式 */
export type XlsxCell = XlsxCellValue | XlsxRichText[] | XlsxCellObject;

/**
 * 一行的两种写法：
 * - 数组式：['A', 'B']，列由下标决定（第 N 项落在第 N 列）
 * - 对象式：{ 1: 'A', C3: 'B' }，键为列号（1 起）或 A1 记号
 */
export type XlsxRowInput = XlsxCell[] | Record<string, XlsxCell>;

/** 解析得到的标准行 */
export interface XlsxRow {
    /** 行号（1 起） */
    index: number;
    /** 该行单元格，数组下标 + 1 = 列号 */
    cells: (XlsxCellObject | null)[];
    /** 行高（磅） */
    height?: number;
    /** 是否隐藏 */
    hidden?: boolean;
}

/** 列定义 */
export interface XlsxColumn {
    /** 列宽（字符数） */
    width?: number;
    /** 是否隐藏 */
    hidden?: boolean;
    /** 表头文本 */
    header?: string;
}

/** 工作表 */
export interface XlsxSheet {
    /** 表名 */
    name: string;
    /** 数据行；解析结果（XlsxRow）可直接回传给 jsonToXlsx */
    rows: (XlsxRowInput | XlsxRow)[];
    /** 列定义 */
    columns?: XlsxColumn[];
    /** 合并区域，如 ['A1:C1'] */
    merges?: string[];
    /** 冻结窗格 */
    frozen?: { rows?: number; columns?: number };
    /** 是否隐藏 */
    hidden?: boolean;
    /** 页面方向 */
    orientation?: 'portrait' | 'landscape';
}

/** 工作簿元数据 */
export interface XlsxMetadata {
    title?: string;
    author?: string;
    creator?: string;
    company?: string;
    subject?: string;
    description?: string;
    keywords?: string;
    category?: string;
    created?: string;
    modified?: string;
    lastModifiedBy?: string;
}

/** 完整工作簿 */
export interface XlsxWorkbook {
    metadata?: XlsxMetadata;
    sheets: XlsxSheet[];
    /**
     * 工作簿内的二进制资源（图片等）
     *
     * 单元格通过 `image` 字段引用（`{ assetId }` 或内联数据），
     * 每项带 `anchor` 指明它落在哪张表的哪个单元格。
     */
    assets?: DocAsset[];
}

/** 解析选项 */
export interface XlsxParseOptions {
    /**
     * 二进制资源（单元格图片）处理方式，默认 `inline`
     *
     * - `inline`：单元格里放 data URL
     * - `extract`：字节进 `workbook.assets`，单元格只留 `{ assetId }` 引用
     * - `both`：两者都给
     * - `ignore`：丢弃（默认不解析图片时保持原有行为）
     */
    assets?: DocAssetMode;
    /** 指定要解析的表名；缺省解析全部 */
    sheets?: string[];
    /** 是否把超链接作为单元格属性返回，默认 true */
    includeHyperlinks?: boolean;
    /** 是否把单元格批注返回，默认 false（体积较大） */
    includeNotes?: boolean;
    /** 是否保留富文本分段，默认 true */
    includeRichText?: boolean;
    /** 跳过样式解析（更快、体积更小） */
    skipStyles?: boolean;
}

/** 解析结果 */
export interface XlsxParseResult {
    /** 工作簿 */
    workbook: XlsxWorkbook;
    /** 解析提示 */
    messages: string[];
}

/** 生成选项 */
export interface XlsxGenerateOptions {
    /** 文档创建者 */
    creator?: string;
    /** 创建时间（ISO 字符串），默认当前时间 */
    created?: string;
    /** 是否压缩输出，默认 true */
    zip?: boolean;
    /**
     * 可用资源表：用于解析单元格上的 `{ assetId }` 图片引用
     *
     * 「解析 → 改 → 生成」时把 `xlsxToJson` 返回的 `assets` 原样传回即可。
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
}
