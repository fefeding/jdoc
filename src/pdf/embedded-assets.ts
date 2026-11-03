/**
 * 从「内嵌资源的 pdf worker」中解出 cMap / 标准字体 / wasm，供 Node 端解析使用。
 *
 * 背景：jdoc 的构建脚本（vscode/esbuild.js）会把 pdfjs-dist 的 cmaps /
 * standard_fonts / wasm 以 base64 内嵌进 pdf worker 文件——文件头带有
 * `/* JDOC_PDF_EMBEDDED_ASSETS *​/` 标记和 `globalThis.__JDC_PDF_ASSETS__` 字典。
 * webview 端靠 worker 内的 fetch 拦截器直接用内存资源；而 Node 端（VS Code 扩展
 * 宿主等打包场景）没有 node_modules 可供探测，pdfjs 的 NodeBinaryDataFactory 是
 * 用 fs.readFile 直读 cMapUrl 路径的，资源必须落在磁盘上。这里把字典解包到系统
 * 临时目录（按 worker 内容哈希做缓存目录名，解包一次后跳过）。
 *
 * @module pdf/embedded-assets
 */

/** 解包得到的资源目录（路径以 / 结尾，可直接拼 pdfjs 要求的文件名） */
export interface EmbeddedPdfAssets {
    /** cMap 目录（中文等 CID 字体解码必需） */
    cmaps?: string;
    /** 标准字体目录（未内嵌字体的 Helvetica 等 14 种） */
    standard_fonts?: string;
    /** wasm 目录（JPEG2000 / JBIG2 图像解码） */
    wasm?: string;
}

/** worker 文件头标记，与 esbuild.js 的 EMBEDDED_WORKER_MARKER 一致 */
const MARKER = '/* JDOC_PDF_EMBEDDED_ASSETS */';
/** 内嵌字典的赋值前缀，字典本体是压缩 JSON（不含分号），到第一个 `;` 为止 */
const DICT_PREFIX = 'globalThis.__JDC_PDF_ASSETS__=';

let cached: Promise<EmbeddedPdfAssets | undefined> | undefined;

/**
 * 从内嵌资源的 worker 文件解包资源，返回各资源目录。
 *
 * 非 Node 环境、文件不存在或未内嵌资源时返回 undefined；结果按 worker 内容
 * 哈希缓存在临时目录，重复调用直接命中，不会反复写盘。
 *
 * @param workerPath - 内嵌资源的 pdf worker 文件路径
 */
export function extractEmbeddedPdfAssets(workerPath: string): Promise<EmbeddedPdfAssets | undefined> {
    const isNode = typeof process !== 'undefined' && !!process.versions?.node;
    if (!isNode) return Promise.resolve(undefined);
    // 单宿主进程里 worker 文件是固定的，统一缓存；失败也不反复重试
    cached ||= doExtract(workerPath).catch(() => undefined);
    return cached;
}

async function doExtract(workerPath: string): Promise<EmbeddedPdfAssets | undefined> {
    // 刻意用变量承接 specifier 再动态 import：浏览器构建时打包器不会尝试解析 node: 内置模块
    const fsSpecifier = 'node:fs/promises';
    const pathSpecifier = 'node:path';
    const osSpecifier = 'node:os';
    const cryptoSpecifier = 'node:crypto';
    const fs = (await import(/* @vite-ignore */ fsSpecifier)) as {
        readFile: (p: string) => Promise<Buffer>;
        access: (p: string) => Promise<void>;
        mkdir: (p: string, o?: { recursive?: boolean }) => Promise<void>;
        writeFile: (p: string, d: string | Buffer) => Promise<void>;
        rm: (p: string, o?: { recursive?: boolean; force?: boolean }) => Promise<void>;
    };
    const path = (await import(/* @vite-ignore */ pathSpecifier)) as {
        join: (...a: string[]) => string;
        dirname: (p: string) => string;
    };
    const os = (await import(/* @vite-ignore */ osSpecifier)) as { tmpdir: () => string };
    const { createHash } = (await import(/* @vite-ignore */ cryptoSpecifier)) as {
        createHash: (algo: string) => {
            update: (d: Buffer) => { digest: (enc: string) => string };
        };
    };

    let raw: Buffer;
    try {
        raw = await fs.readFile(workerPath);
    } catch {
        return undefined;
    }
    if (raw.length < MARKER.length || !raw.subarray(0, MARKER.length).equals(Buffer.from(MARKER))) {
        return undefined;
    }
    const text = raw.toString('utf8');
    const start = text.indexOf(DICT_PREFIX);
    if (start < 0) return undefined;
    const from = start + DICT_PREFIX.length;
    const end = text.indexOf(';', from);
    if (end < 0) return undefined;
    let dict: Record<string, string>;
    try {
        dict = JSON.parse(text.slice(from, end)) as Record<string, string>;
    } catch {
        return undefined;
    }
    const keys = Object.keys(dict);
    if (!keys.length) return undefined;

    // 目录名带内容哈希：worker 更新（pdfjs 升级 / 资源变化）后自动换新目录
    const root = path.join(os.tmpdir(), `jdoc-pdf-assets-${createHash('sha1').update(raw).digest('hex').slice(0, 12)}`);
    try {
        await fs.access(path.join(root, '.complete'));
        return toAssets(root, keys);
    } catch {
        // 未解包过（或上次中断），继续
    }
    await fs.rm(root, { recursive: true, force: true });
    await fs.mkdir(root, { recursive: true });
    for (const [name, b64] of Object.entries(dict)) {
        // name 形如 cmaps/UniGB-UCS2-H.bcmap
        const file = path.join(root, name);
        await fs.mkdir(path.dirname(file), { recursive: true });
        await fs.writeFile(file, Buffer.from(b64, 'base64'));
    }
    // 完成标记：解包中断（进程被杀）不会留下假缓存
    await fs.writeFile(path.join(root, '.complete'), '');
    return toAssets(root, keys);
}

/** 只暴露字典里真实存在的子目录，避免把 pdfjs 指向不存在的目录 */
function toAssets(root: string, keys: string[]): EmbeddedPdfAssets {
    const out: EmbeddedPdfAssets = {};
    for (const sub of ['cmaps', 'standard_fonts', 'wasm'] as const) {
        if (keys.some((k) => k.startsWith(`${sub}/`))) out[sub] = `${root}/${sub}/`;
    }
    return out;
}

/**
 * 宿主侧（Node / Electron 扩展宿主）专用的 pdfjs 二进制资源工厂。
 *
 * pdfjs 在 Electron 里会把 `isNodeJS` 判成 false（因为 `process.versions.electron`
 * 存在），于是选 `DOMBinaryDataFactory` 去读 `document.baseURI` —— 而扩展宿主没有
 * `document`，直接抛 `ReferenceError: document is not defined`。这里提供用 `fs` 读磁盘
 * 资源的工厂，绕过 DOM 分支：只要把 `getDocument({ BinaryDataFactory: FsBinaryDataFactory })`
 * 传进去，pdfjs 就不会再触碰 `document`。
 *
 * 注意：仅 Node 端使用（parsePdf 里 `isNode` 判断）。浏览器端 pdfjs 走 DOM 工厂 +
 * 内嵌 worker 的 fetch 拦截即可，不需要这个。接口对齐 pdfjs 的 BaseBinaryDataFactory：
 * `fetch({ kind, filename })` 返回 Uint8Array，kind 为 cMapUrl / standardFontDataUrl / wasmUrl。
 */
export class FsBinaryDataFactory {
    cMapUrl: string | null;
    standardFontDataUrl: string | null;
    wasmUrl: string | null;
    constructor({
        cMapUrl = null,
        standardFontDataUrl = null,
        wasmUrl = null
    }: {
        cMapUrl?: string | null;
        standardFontDataUrl?: string | null;
        wasmUrl?: string | null;
    }) {
        this.cMapUrl = cMapUrl ?? null;
        this.standardFontDataUrl = standardFontDataUrl ?? null;
        this.wasmUrl = wasmUrl ?? null;
    }
    async fetch({ kind, filename }: { kind: string; filename: string }): Promise<Uint8Array> {
        let base: string | null | undefined;
        switch (kind) {
            case 'cMapUrl':
                base = this.cMapUrl;
                break;
            case 'standardFontDataUrl':
                base = this.standardFontDataUrl;
                break;
            case 'wasmUrl':
                base = this.wasmUrl;
                break;
            default:
                throw new Error(`Not implemented: ${kind}`);
        }
        if (!base) throw new Error(`Ensure that the \`${kind}\` API parameter is provided.`);
        const url = `${base}${filename}`;
        const fs = (await import(/* @vite-ignore */ 'node:fs/promises')) as {
            readFile: (p: string) => Promise<Buffer>;
        };
        try {
            const buf = await fs.readFile(url);
            return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
        } catch {
            throw new Error(`Unable to load ${kind} data at: ${url}`);
        }
    }
}
