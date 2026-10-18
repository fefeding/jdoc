/**
 * 纯文本（TXT / LOG 等）解析
 *
 * 仅做 UTF-8 解码 + HTML 转义，供预览使用；不做任何标记解析。
 *
 * @module text
 */

import { wrapError } from '../core/errors';
import { toUint8Array, type DocInput } from '../core/input';

/** textToHtml 的返回结构 */
export interface TextHtmlResult {
    /** 渲染后的 HTML（根元素带 `jdoc-txt` 类） */
    html: string;
    /** 原文纯文本 */
    text: string;
    /** 解析提示 */
    messages: string[];
}

/** textToHtml 的可选项 */
export interface TextOptions {
    /** 是否附带默认样式，默认 true */
    includeStyle?: boolean;
}

const TEXT_CSS = `
.jdoc-txt { margin: 0; font: 13px/1.7 ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace;
  white-space: pre-wrap; word-break: break-word; color: var(--jdoc-md-fg, #1f2328); }
`;

/** HTML 转义 */
function escapeHtml(value: string): string {
    return value.replace(/[&<>"']/g, (c) => {
        switch (c) {
            case '&': return '&amp;';
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '"': return '&quot;';
            default: return '&#39;';
        }
    });
}

/** 字节 -> UTF-8 文本（去 BOM） */
function decodeText(bytes: Uint8Array): string {
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
        bytes = bytes.subarray(3);
    }
    return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
}

/**
 * 渲染纯文本为 HTML
 * @param input - Uint8Array / ArrayBuffer / Blob / base64 字符串
 */
export async function textToHtml(input: DocInput, options: TextOptions = {}): Promise<TextHtmlResult> {
    try {
        const text = decodeText(await toUint8Array(input));
        const style = options.includeStyle === false ? '' : `<style>${TEXT_CSS}</style>`;
        return {
            html: `${style}<pre class="jdoc-txt">${escapeHtml(text)}</pre>`,
            text,
            messages: []
        };
    } catch (error) {
        throw wrapError('txt', '解析', error);
    }
}

/**
 * 提取纯文本
 * @param input - 同 textToHtml
 */
export async function textToText(input: DocInput): Promise<string> {
    try {
        return decodeText(await toUint8Array(input));
    } catch (error) {
        throw wrapError('txt', '解析', error);
    }
}
