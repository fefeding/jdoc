/**
 * 「转换格式」命令的格式映射
 *
 * 单独成文件是为了让宿主逻辑（extension.ts）与回归脚本（scripts/verify-convert-formats.mjs）
 * 共用同一份映射，避免两边各写一份扩展名正则而漂移。
 *
 * 该文件不依赖 vscode API，可直接在 Node 脚本里引入。
 */

import type { ConvertTarget } from '@fefeding/jdoc';

/**
 * 转换命令支持的源格式（与文件扩展名对应）
 *
 * 'xlsx' 走库里的 xlsxToBlocks（工作表 -> 标题 + 表格），因此 Excel 也能导出成
 * Word / PDF / Markdown / HTML。
 * 'pptx' 走库里的 pptxToBlocks（幻灯片 -> 标题 + 正文块），因此 PPT 也能导出成
 * Word / PDF / Markdown / HTML。
 */
export type SourceFormat = 'docx' | 'pdf' | 'md' | 'html' | 'txt' | 'xlsx' | 'pptx';

/** 文件扩展名 -> 源格式（传给 convert 的 from 参数） */
export function extToFormat(path: string): SourceFormat | undefined {
    if (/\.docx$/i.test(path)) return 'docx';
    if (/\.pdf$/i.test(path)) return 'pdf';
    // 注意：不能用 /\.markdown?$/ —— 那只匹配「.markdow/.markdown」，漏掉最常见的「.md」
    if (/\.(md|markdown)$/i.test(path)) return 'md';
    if (/\.html?$/i.test(path)) return 'html';
    if (/\.txt$/i.test(path)) return 'txt';
    if (/\.xlsx$/i.test(path)) return 'xlsx';
    if (/\.pptx$/i.test(path)) return 'pptx';
    return undefined;
}

/** 源格式 -> 可选目标格式（去掉与源相同的格式） */
export function targetFormats(from: SourceFormat, targets: readonly ConvertTarget[]): ConvertTarget[] {
    return targets.filter((item) => item !== from);
}
