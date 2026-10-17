/**
 * PPTX 解析与生成公开入口
 *
 * PPTX 能力由 @fefeding/ppt-parser 提供，本模块负责：
 * 1) 把 jdoc 的宽松入参（Blob / base64 / TypedArray）转成底层库要求的 ArrayBuffer
 * 2) 把底层库的宽松出参（string | Buffer | Blob | ...）统一为 Uint8Array
 * 3) 导出标准文档契约 PptxDocument，供 jsonToPptx 往返
 *
 * @module pptx
 */

import {
    PPTXComposer as BaseComposer,
    editPptx as baseEditPptx,
    jsonToPptx as baseJsonToPptx,
    pptxToFiles as basePptxToFiles,
    pptxToHtml as basePptxToHtml,
    pptxToJson as basePptxToJson,
    pptxToStandard as basePptxToStandard
} from '@fefeding/ppt-parser';
import { wrapError } from '../core/errors';
import { toArrayBuffer, type DocInput } from '../core/input';

export type { PptxDocument } from '@fefeding/ppt-parser';
export { PPTXShapeUtils, presetShapePath, createStore, createActions, renderChartSVG } from '@fefeding/ppt-parser';
export { normalizeDoc, normalizeElement } from '@fefeding/ppt-parser';

/** 底层库可能返回多种类型，这里统一收敛为 Uint8Array */
function toUint8(result: unknown): Uint8Array {
    if (result instanceof Uint8Array) return result;
    if (typeof result === 'string') {
        const binary = (globalThis as { atob?: (s: string) => string }).atob
            ? (globalThis as { atob: (s: string) => string }).atob(result)
            : Buffer.from(result, 'base64').toString('binary');
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i) & 0xff;
        return bytes;
    }
    if (result instanceof ArrayBuffer) return new Uint8Array(result);
    if (Array.isArray(result)) return new Uint8Array(result);
    throw new TypeError('PPTX 生成结果类型无法识别');
}

/** PPTX 生成选项 */
export interface PptxGenerateOptions {
    [key: string]: unknown;
}

/**
 * PPTX 解析结果的宽松类型。
 * ppt-parser 的返回结构引用了其包内未导出的类型（HtmlSlideResult / XmlNode / PptxMetadata 等），
 * 直接透传会让声明文件无法生成，因此对外以宽松类型暴露，运行时结构保持不变。
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type PptxLooseResult = any;

/**
 * 解析 PPTX 为 HTML 幻灯片
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 透传 ppt-parser 的解析选项
 */
export async function pptxToHtml(
    input: DocInput,
    options: Record<string, unknown> = {}
): Promise<PptxLooseResult> {
    try {
        return await basePptxToHtml(await toArrayBuffer(input), options as never);
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/**
 * 解析 PPTX 为结构化 JSON
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 透传 ppt-parser 的解析选项
 */
export async function pptxToJson(
    input: DocInput,
    options: Record<string, unknown> = {}
): Promise<PptxLooseResult> {
    try {
        return await basePptxToJson(await toArrayBuffer(input), options as never);
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/**
 * 解析 PPTX 为标准文档结构（与 jsonToPptx 同源，可直接往返）
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 透传 ppt-parser 的解析选项
 */
export async function pptxToStandard(
    input: DocInput,
    options: Record<string, unknown> = {}
): Promise<PptxLooseResult> {
    try {
        return await basePptxToStandard(await toArrayBuffer(input), options as never);
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/** 列出 PPTX 内部文件及内容 */
export async function pptxToFiles(input: DocInput): Promise<PptxLooseResult> {
    try {
        return await basePptxToFiles(await toArrayBuffer(input));
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}

/**
 * 生成 PPTX 文件
 * @param presentation - 标准文档结构（或 ppt-parser 兼容的 JSON）
 * @returns PPTX 字节数据
 */
export async function jsonToPptx(
    presentation: unknown,
    options: PptxGenerateOptions = {}
): Promise<Uint8Array> {
    try {
        const result = await baseJsonToPptx(presentation, { ...options, outputType: 'uint8array' });
        return toUint8(result);
    } catch (error) {
        throw wrapError('pptx', '生成', error);
    }
}

/** 链式构建 PPTX（save() 统一返回 Uint8Array） */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const PPTXComposer = BaseComposer as unknown as new (
    options?: Record<string, unknown>
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
) => any;

/**
 * 打开已有 PPTX 进行编辑（增删页、改元数据等）
 * @param input - 文件二进制
 */
export async function editPptx(input: DocInput): Promise<PptxLooseResult> {
    try {
        return await baseEditPptx(await toArrayBuffer(input));
    } catch (error) {
        throw wrapError('pptx', '解析', error);
    }
}
