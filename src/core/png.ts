/**
 * 极简 PNG 编码器
 *
 * 场景：PDF 解析时 pdfjs 给出的是**解码后的原始像素**（灰度 / RGB / RGBA），
 * 必须编码成 PNG 才能当图片资源用（写进 docx / html / md）。
 *
 * 这里刻意不引入压缩依赖，按可用性三级降级：
 * 1. 浏览器 / Worker：`CompressionStream('deflate')`（产出 RFC1950 zlib 容器，正是 PNG 需要的）
 * 2. Node：`node:zlib` 的 `deflateSync`
 * 3. 都没有：手写「未压缩 stored 块」的 deflate 流（体积大，但仍是合法 PNG）
 *
 * @module core/png
 */

/** 原始位图（兼容 pdfjs 的 ImageKind） */
export interface RawBitmap {
    /** 像素数据（行优先，从上到下） */
    data: Uint8Array;
    /** 宽度（像素） */
    width: number;
    /** 高度（像素） */
    height: number;
    /** pdfjs ImageKind：1=灰度 1bpp，2=RGB 24bpp，3=RGBA 32bpp；缺省按数据长度推断 */
    kind?: number;
}

/** PNG 色彩类型 */
const COLOR_GRAY = 0;
const COLOR_RGB = 2;
const COLOR_RGBA = 6;

/** PNG 文件头 */
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/**
 * 把原始位图编码为 PNG 字节
 * @param bitmap - 原始像素（pdfjs 解出的灰度 / RGB / RGBA）
 * @returns PNG 文件字节
 */
export async function encodePng(bitmap: RawBitmap): Promise<Uint8Array> {
    const { data, width, height } = bitmap;
    if (!width || !height || !data?.length) throw new TypeError('位图尺寸或像素数据为空');

    const { channels, colorType } = resolveFormat(bitmap);
    const stride = width * channels;
    // 每行前置 1 字节 filter 类型，并按行自适应选滤波器：
    // 全用 None（filter 0）时照片类内容几乎压不动，体积会是自适应滤波的数倍。
    const raw = filterRows(bitmap, height, stride, channels);

    const idat = await deflateZlib(raw);
    const ihdr = new Uint8Array(13);
    const view = new DataView(ihdr.buffer);
    view.setUint32(0, width);
    view.setUint32(4, height);
    ihdr[8] = 8; // bitDepth
    ihdr[9] = colorType;
    ihdr[10] = 0; // compression: deflate
    ihdr[11] = 0; // filter: adaptive
    ihdr[12] = 0; // interlace: none

    return concat([
        new Uint8Array(PNG_SIGNATURE),
        chunk('IHDR', ihdr),
        chunk('IDAT', idat),
        chunk('IEND', new Uint8Array(0))
    ]);
}

/** PNG 行滤波器类型（0=None 1=Sub 2=Up 3=Average 4=Paeth） */
const FILTERS = [0, 1, 2, 3, 4];

/**
 * 逐行做滤波并自适应选类型（最小绝对差之和）
 *
 * @param bitmap - 原始位图
 * @param height - 行数
 * @param stride - 每行字节数
 * @param channels - 通道数（Sub 滤波的「左邻」偏移）
 */
function filterRows(bitmap: RawBitmap, height: number, stride: number, channels: number): Uint8Array {
    const raw = new Uint8Array((stride + 1) * height);
    const prev = new Uint8Array(stride); // 上一行原始像素（Up 滤波参考）
    const best = new Uint8Array(stride);
    const candidate = new Uint8Array(stride);
    for (let y = 0; y < height; y++) {
        const src = normalizeRow(bitmap, y, stride);
        let bestType = 0;
        let bestScore = Infinity;
        for (const type of FILTERS) {
            applyFilter(type, src, prev, stride, channels, candidate);
            const score = sumAbs(candidate);
            if (score < bestScore) {
                bestScore = score;
                bestType = type;
                best.set(candidate);
            }
        }
        const dst = y * (stride + 1);
        raw[dst] = bestType;
        raw.set(best, dst + 1);
        prev.set(src);
    }
    return raw;
}

/** 按类型对一行做滤波（结果写入 out，所有运算按 256 取模） */
function applyFilter(
    type: number,
    src: Uint8Array,
    prev: Uint8Array,
    stride: number,
    channels: number,
    out: Uint8Array
): void {
    if (type === 0) {
        out.set(src);
        return;
    }
    for (let i = 0; i < stride; i++) {
        const left = i >= channels ? src[i - channels] : 0;
        const up = prev[i];
        const upLeft = i >= channels ? prev[i - channels] : 0;
        let value: number;
        switch (type) {
            case 1:
                value = src[i] - left;
                break;
            case 2:
                value = src[i] - up;
                break;
            case 3:
                value = src[i] - ((left + up) >> 1);
                break;
            default: {
                const p = left + up - upLeft;
                const pa = Math.abs(p - left);
                const pb = Math.abs(p - up);
                const pc = Math.abs(p - upLeft);
                value = src[i] - (pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft);
                break;
            }
        }
        out[i] = value & 0xff;
    }
}

/** 滤波后行的「绝对差之和」（按有符号字节计，越小越易压缩） */
function sumAbs(row: Uint8Array): number {
    let sum = 0;
    for (let i = 0; i < row.length; i++) sum += row[i] < 128 ? row[i] : 256 - row[i];
    return sum;
}

/** 依据 ImageKind / 数据长度判断通道数与 PNG 色彩类型 */
function resolveFormat(bitmap: RawBitmap): { channels: number; colorType: number } {
    const { data, width, height, kind } = bitmap;
    if (kind === 1) return { channels: 1, colorType: COLOR_GRAY };
    if (kind === 2) return { channels: 3, colorType: COLOR_RGB };
    if (kind === 3) return { channels: 4, colorType: COLOR_RGBA };
    const perPixel = data.length / Math.max(1, width * height);
    if (perPixel >= 4) return { channels: 4, colorType: COLOR_RGBA };
    if (perPixel >= 3) return { channels: 3, colorType: COLOR_RGB };
    return { channels: 1, colorType: COLOR_GRAY };
}

/** 取第 y 行归一化后的字节（1bpp 灰度展开成 8bpp，其余原样） */
function normalizeRow(bitmap: RawBitmap, y: number, stride: number): Uint8Array {
    const { data, width, height, kind } = bitmap;
    const row = new Uint8Array(stride);
    if (kind === 1) {
        // 1bpp：每字节 8 个像素，MSB 在先；位为 1 表示「着色」（黑），0 表示不着色（白）
        const srcStride = Math.ceil(width / 8);
        for (let x = 0; x < width; x++) {
            const byte = data[y * srcStride + (x >> 3)] ?? 0;
            const bit = (byte >> (7 - (x & 7))) & 1;
            row[x] = bit ? 0 : 255;
        }
        return row;
    }
    const src = y * stride;
    const available = Math.min(stride, Math.max(0, data.length - src));
    if (available > 0) row.set(data.subarray(src, src + available), 0);
    // 数据不足时（异常位图）剩余补 0，避免整行错位
    void height;
    return row;
}

/** 拼接多个字节段 */
function concat(parts: Uint8Array[]): Uint8Array {
    const total = parts.reduce((sum, p) => sum + p.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    for (const part of parts) {
        out.set(part, offset);
        offset += part.length;
    }
    return out;
}

/** 组装一个 PNG chunk（长度 + 类型 + 数据 + CRC32） */
function chunk(type: string, data: Uint8Array): Uint8Array {
    const out = new Uint8Array(12 + data.length);
    const view = new DataView(out.buffer);
    view.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
}

// ---------------------------------------------------------------------------
// CRC32 / Adler32
// ---------------------------------------------------------------------------

/** CRC32 表（PNG chunk 校验用） */
const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

function crc32(bytes: Uint8Array): number {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
}

function adler32(bytes: Uint8Array): number {
    let a = 1;
    let b = 0;
    for (let i = 0; i < bytes.length; i++) {
        a = (a + bytes[i]) % 65521;
        b = (b + a) % 65521;
    }
    return ((b << 16) | a) >>> 0;
}

// ---------------------------------------------------------------------------
// deflate（zlib 容器）
// ---------------------------------------------------------------------------

/** 对原始字节做 zlib 压缩（PNG 的 IDAT 需要 RFC1950 容器） */
async function deflateZlib(raw: Uint8Array): Promise<Uint8Array> {
    // 1) 浏览器 / Worker：CompressionStream('deflate') 直接产出 zlib 容器
    if (typeof CompressionStream === 'function' && typeof Response === 'function') {
        try {
            const stream = new CompressionStream('deflate');
            const writer = stream.writable.getWriter() as unknown as {
                write(chunk: unknown): Promise<void>;
                close(): Promise<void>;
            };
            const pending = new Response(stream.readable).arrayBuffer();
            await writer.write(raw);
            await writer.close();
            return new Uint8Array(await pending);
        } catch {
            // 某些宿主实现不完整，继续走后面的兜底
        }
    }

    // 2) Node：zlib.deflateSync 默认输出 zlib 容器
    if (typeof process !== 'undefined' && !!process.versions?.node) {
        try {
            const specifier = 'node:zlib';
            const mod = (await import(/* @vite-ignore */ specifier)) as {
                deflateSync?: (data: Uint8Array) => Uint8Array;
            };
            if (typeof mod.deflateSync === 'function') return new Uint8Array(mod.deflateSync(raw));
        } catch {
            // 继续走兜底
        }
    }

    // 3) 兜底：未压缩的 stored deflate 块（合法但体积≈原始数据）
    return storedDeflate(raw);
}

/** 手写 stored（未压缩）deflate 块 + zlib 头尾 */
function storedDeflate(raw: Uint8Array): Uint8Array {
    const MAX_BLOCK = 0xffff;
    const blocks: number[] = [0x78, 0x01]; // zlib: CMF=0x78, FLG=0x01（最低压缩等级）
    let offset = 0;
    while (offset < raw.length) {
        const size = Math.min(MAX_BLOCK, raw.length - offset);
        const final = offset + size >= raw.length ? 1 : 0;
        blocks.push(final); // BFINAL(bit0) + BTYPE=00(stored)
        blocks.push(size & 0xff, (size >> 8) & 0xff);
        blocks.push(~size & 0xff, (~size >> 8) & 0xff);
        for (let i = 0; i < size; i++) blocks.push(raw[offset + i]);
        offset += size;
    }
    const adler = adler32(raw);
    blocks.push((adler >> 24) & 0xff, (adler >> 16) & 0xff, (adler >> 8) & 0xff, adler & 0xff);
    return new Uint8Array(blocks);
}
