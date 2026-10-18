/**
 * jdoc — 文档解析与生成统一库
 *
 * 支持 PPTX / DOCX / XLSX / PDF 四种格式的解析与生成，浏览器与 Node.js 双端可用。
 * 各格式提供独立入口，命名风格统一为：
 *   解析 xxxToYyy(input, options)   生成 jsonToXxx(doc, options)
 *
 * @module jdoc
 */

import type { DocFormat } from './core/errors';
import { toUint8Array, type DocInput } from './core/input';

export * from './core/errors';
export * from './core/input';
export * from './core/xml';
export * from './core/asset';
export * from './core/bundle';

export * from './pptx/index';
export * from './docx/index';
export * from './xlsx/index';
export * from './pdf/index';
export * from './markdown/index';
export * from './text/index';

/** Markdown 强特征：出现任意一个即可判定为 Markdown */
const MD_STRONG = [
    /^#{1,6}\s+\S/m,            // 标题
    /^```/m,                    // 代码围栏
    /\[[^\]\n]+\]\([^)\n]+\)/,  // 链接
    /!\[[^\]\n]*\]\([^)\n]+\)/, // 图片
    /^\s*\|.*\|\s*$/m           // 表格行
];

/** Markdown 弱特征：单看不足为证（普通文本也常见），需同时命中两个 */
const MD_WEAK = [
    /^\s*[-*+]\s+\S/m,          // 无序列表
    /^\s*\d+\.\s+\S/m,          // 有序列表
    /^>\s+\S/m,                 // 引用
    /^\s*(?:-{3,}|\*{3,}|_{3,})\s*$/m // 分隔线
];

/** 判断一段文本是否像 Markdown（强特征命中 1 个，或弱特征命中 2 个） */
function looksLikeMarkdown(text: string): boolean {
    if (MD_STRONG.some((re) => re.test(text))) return true;
    return MD_WEAK.filter((re) => re.test(text)).length >= 2;
}

/**
 * 识别文件格式（按魔数与 ZIP 内部结构判断）
 * @param input - 文件二进制或 data URL / base64 字符串
 * @returns 识别到的格式；无法识别返回 null
 *
 * @example
 * ```js
 * const type = await detectFormat(file); // 'docx' | 'xlsx' | 'pptx' | 'pdf' | null
 * ```
 */
export async function detectFormat(input: DocInput): Promise<DocFormat | null> {
    const bytes = await toUint8Array(input);
    if (!bytes.length) return null;

    if (bytes.length >= 4) {
        // PDF: %PDF-
        if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) return 'pdf';

        // ZIP 容器（OOXML）：PK\x03\x04 / PK\x05\x06 / PK\x07\x08
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
            const isZip =
                (bytes[2] === 0x03 && bytes[3] === 0x04) ||
                (bytes[2] === 0x05 && bytes[3] === 0x06) ||
                (bytes[2] === 0x07 && bytes[3] === 0x08);
            if (!isZip) return null;

            // OOXML 各格式靠内部标记部件区分：
            //   docx -> word/document.xml   xlsx -> xl/workbook.xml   pptx -> ppt/presentation.xml
            const text = await new TextDecoder('latin1').decode(bytes.slice(0, Math.min(bytes.length, 8192)));
            if (text.includes('word/')) return 'docx';
            if (text.includes('xl/')) return 'xlsx';
            if (text.includes('ppt/')) return 'pptx';

            // 中央目录可能在文件末尾，尾部再找一次
            const tail = await new TextDecoder('latin1').decode(bytes.slice(Math.max(0, bytes.length - 65536)));
            if (tail.includes('word/')) return 'docx';
            if (tail.includes('xl/')) return 'xlsx';
            if (tail.includes('ppt/')) return 'pptx';
            return null;
        }
    }

    // 文本类格式（Markdown / 纯文本）：无二进制魔数，按"可 UTF-8 解码 + 无 NUL 字节"判定。
    // 放在魔数判断之后，保证 %PDF- / PK 开头的二进制文件优先按原格式识别。
    if (!bytes.includes(0)) {
        try {
            const head = new TextDecoder('utf-8', { fatal: true }).decode(bytes.slice(0, Math.min(bytes.length, 8192)));
            return looksLikeMarkdown(head) ? 'md' : 'txt';
        } catch {
            // 不是合法 UTF-8，无法判定
        }
    }

    return null;
}

/** jdoc 版本号（与 package.json 保持一致） */
export const VERSION = '0.1.0';
