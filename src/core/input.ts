/**
 * 输入归一化模块
 *
 * 四种格式的底层库（mammoth / exceljs / pdfjs-dist / ppt-parser）对入参类型要求各不相同：
 * 有的要 ArrayBuffer，有的要 Buffer，有的要 { arrayBuffer() }。本模块把这些差异全部收口，
 * 让调用方无论传入 ArrayBuffer / TypedArray / Buffer / Blob / File / data URL 字符串都能直接用。
 *
 * @module core/input
 */

/** 统一的文件输入类型：二进制缓冲区、Blob/File、或 data URL / base64 字符串 */
export type DocInput = ArrayBuffer | ArrayBufferView | Uint8Array | Blob | string;

/** 去掉 data URL 前缀，提取纯 base64 负载 */
function stripDataUrlPrefix(value: string): string {
    const match = /^data:[^,]*?(;base64)?,(.*)$/s.exec(value.trim());
    if (match) {
        // 非 base64 的 data URL（如 data:text/plain,hello）按 UTF-8 文本处理
        if (!match[1]) return '';
        return match[2];
    }
    return value.trim();
}

/** 判断是否为 Node.js Buffer（Buffer 是 Uint8Array 子类，需在 ArrayBufferView 之前判断） */
function isNodeBuffer(value: unknown): boolean {
    return typeof Buffer !== 'undefined' && Buffer.isBuffer(value);
}

/** base64 字符串 -> Uint8Array（双端可用，不依赖 atob / Buffer） */
function base64ToBytes(base64: string): Uint8Array {
    const clean = base64.replace(/\s+/g, '');
    if (!clean) return new Uint8Array(0);

    const globalAtob = (globalThis as { atob?: (s: string) => string }).atob;
    const binary = globalAtob
        ? globalAtob(clean)
        : Buffer.from(clean, 'base64').toString('binary');

    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i) & 0xff;
    }
    return bytes;
}

/** Uint8Array -> base64（双端可用） */
export function bytesToBase64(bytes: Uint8Array): string {
    const globalBtoa = (globalThis as { btoa?: (s: string) => string }).btoa;
    if (globalBtoa) {
        let binary = '';
        // 分块避免超长数组触发 apply 参数上限
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunkSize)) as unknown as number[]);
        }
        return globalBtoa(binary);
    }
    return Buffer.from(bytes).toString('base64');
}

/**
 * 把任意支持的输入统一转换为 Uint8Array
 * @param input - ArrayBuffer / TypedArray / Uint8Array / Blob / File / data URL / base64 字符串
 * @returns 字节数组视图（不复制底层内存时与入参共享内存）
 */
export async function toUint8Array(input: DocInput): Promise<Uint8Array> {
    if (input == null) {
        throw new TypeError('jdoc: 缺少输入数据');
    }

    if (typeof input === 'string') {
        return base64ToBytes(stripDataUrlPrefix(input));
    }

    if (isNodeBuffer(input)) {
        // Buffer 可能来自更大的共享 ArrayBuffer，务必按自身偏移切片
        const buf = input as Buffer;
        return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    }

    if (input instanceof Uint8Array) {
        return input;
    }

    if (ArrayBuffer.isView(input)) {
        return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    }

    if (input instanceof ArrayBuffer) {
        return new Uint8Array(input);
    }

    // Blob / File：Node 18+ 与浏览器均提供 arrayBuffer()
    const maybeBlob = input as Blob;
    if (typeof maybeBlob.arrayBuffer === 'function') {
        return new Uint8Array(await maybeBlob.arrayBuffer());
    }

    throw new TypeError('jdoc: 不支持的输入类型，请传入 ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串');
}

/**
 * 把任意支持的输入统一转换为 ArrayBuffer
 * @param input - 同 toUint8Array
 * @returns ArrayBuffer（始终是一块独立可传递的内存）
 */
export async function toArrayBuffer(input: DocInput): Promise<ArrayBuffer> {
    if (typeof input === 'string') {
        const bytes = base64ToBytes(stripDataUrlPrefix(input));
        return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    }

    if (input instanceof ArrayBuffer) {
        return input;
    }

    if (isNodeBuffer(input)) {
        // Buffer 的底层内存常大于其自身长度，必须精确切片，否则会带上相邻内存
        const buf = input as Buffer;
        return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
    }

    const bytes = await toUint8Array(input);
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/**
 * 传给底层库时的通用二进制输入：优先 Buffer（Node 生态库对 Buffer 兼容最好），浏览器回退 ArrayBuffer
 */
export async function toBinaryPayload(input: DocInput): Promise<ArrayBuffer | Uint8Array> {
    const bytes = await toUint8Array(input);
    if (typeof Buffer !== 'undefined') {
        return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    }
    return bytes;
}

/** 校验二进制长度，避免把空文件/误传文件喂给底层库导致难懂的报错 */
export function assertNonEmpty(bytes: Uint8Array | ArrayBuffer, format: string, minSize = 4): void {
    const size = bytes instanceof ArrayBuffer ? bytes.byteLength : bytes.byteLength;
    if (size < minSize) {
        throw new TypeError(`jdoc: ${format} 数据无效（长度 ${size} 字节，至少需要 ${minSize} 字节）`);
    }
}