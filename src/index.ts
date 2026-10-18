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
    if (bytes.length < 4) return null;

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
    }

    return null;
}

/** jdoc 版本号（与 package.json 保持一致） */
export const VERSION = '0.1.0';
