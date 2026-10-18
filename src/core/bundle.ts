/**
 * 文档 + 资源 的「一体保存」打包 / 解包
 *
 * 背景：解析出的文档结构与二进制资源是分离的两份数据。若要「另存为」，
 * 单存 JSON 会丢失字节，单存字节会丢失结构与位置关系。本模块把两者打成
 * 一个自包含的 bundle（ZIP），任意机器解压即可还原：
 *
 * ```
 * report.jdbundle
 * ├── manifest.json        # 格式、名称、版本、资源清单（含位置锚点）
 * ├── document.docx        # 原始文档（若来源是文件）
 * └── assets/
 *     ├── docx-0001.png     # 资源字节，文件名与资源 id 对齐
 *     └── ...
 * ```
 *
 * 打包用 JSZip（已是依赖），双端可用；落盘用 node:fs（仅 Node）。
 *
 * @module core/bundle
 */

import JSZip from 'jszip';
import { toUint8Array, type DocInput } from './input';
import { createAssetId, mimeOf, type AssetAnchor, type DocAsset } from './asset';

/** bundle 清单文件（放在压缩包根目录） */
export const BUNDLE_MANIFEST = 'manifest.json';

/** 当前 bundle 格式版本，便于将来演进 */
export const BUNDLE_VERSION = 1;

/** 资源在 bundle 中的存放目录 */
const ASSET_DIR = 'assets/';

/** 一体化文档包的内容 */
export interface DocumentBundle {
    /** 文档格式：docx / xlsx / pptx / pdf */
    format: string;
    /** 文档名（不含扩展名），如「季度报告」 */
    name: string;
    /** 解析出的文档结构（JSON 形态，可直接回传给 jsonToXxx） */
    document?: unknown;
    /** 资源列表，保留 anchor 位置信息 */
    assets: DocAsset[];
    /** 原始文档字节（另存为时一并保留，避免重新生成） */
    original?: Uint8Array;
    /** 清单里的其它自定义信息 */
    extra?: Record<string, unknown>;
}

/** 资源清单条目（JSON 友好，不含字节） */
interface ManifestAsset {
    id: string;
    name: string;
    mime: string;
    /** bundle 内的相对路径，如 assets/docx-0001.png */
    path: string;
    /** 原始部件路径 */
    part?: string;
    anchor: AssetAnchor;
    width?: number;
    height?: number;
    alt?: string;
    size: number;
}

/** bundle 清单 */
export interface BundleManifest {
    /** 固定为 'jdoc-bundle'，便于识别 */
    kind: 'jdoc-bundle';
    /** 清单结构版本 */
    version: number;
    format: string;
    name: string;
    /** 原始文档在包内的相对路径 */
    original?: string;
    /** document.json 是否存在 */
    hasDocument: boolean;
    assets: ManifestAsset[];
    /** 其它自定义字段 */
    extra?: Record<string, unknown>;
}

/** 由 MIME 补全扩展名 */
function extFor(mime: string, fallbackName: string): string {
    const ext = fallbackName.includes('.')
        ? `.${fallbackName.split('.').pop()}`
        : '';
    const known: Record<string, string> = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/svg+xml': '.svg',
        'application/pdf': '.pdf'
    };
    return known[mime] || ext || '.bin';
}

/**
 * 把文档与资源打成一体包（Uint8Array / ZIP 字节）
 *
 * @param bundle - 文档结构 + 资源 + 可选原始字节
 * @returns ZIP 字节，可直接存盘或下载
 *
 * @example
 * ```js
 * const parsed = await docxToJson(file);
 * const zip = await packDocument({
 *     format: 'docx',
 *     name: '合同',
 *     document: parsed,
 *     assets: parsed.assets,
 * });
 * await writeFile('合同.jdbundle', zip);
 * ```
 */
export async function packDocument(bundle: DocumentBundle): Promise<Uint8Array> {
    const zip = new JSZip();
    const assets = bundle.assets || [];
    const usedNames = new Set<string>();

    const manifestAssets: ManifestAsset[] = [];
    assets.forEach((asset, index) => {
        const id = asset.id || createAssetId(bundle.format, index);
        // 文件名去重，避免同名资源互相覆盖
        let fileName = `${id}${extFor(asset.mime || mimeOf(asset.part || asset.name), asset.name)}`;
        let seq = 1;
        while (usedNames.has(fileName)) {
            fileName = `${id}-${seq++}${extFor(asset.mime, asset.name)}`;
        }
        usedNames.add(fileName);
        const path = `${ASSET_DIR}${fileName}`;
        zip.file(path, asset.data, { binary: true });
        const entry: ManifestAsset = {
            id,
            name: asset.name || fileName,
            mime: asset.mime || 'application/octet-stream',
            path,
            anchor: asset.anchor,
            size: asset.data?.byteLength || 0
        };
        if (asset.part) entry.part = asset.part;
        if (asset.width != null) entry.width = asset.width;
        if (asset.height != null) entry.height = asset.height;
        if (asset.alt) entry.alt = asset.alt;
        manifestAssets.push(entry);
    });

    const manifest: BundleManifest = {
        kind: 'jdoc-bundle',
        version: BUNDLE_VERSION,
        format: bundle.format,
        name: bundle.name || 'document',
        hasDocument: bundle.document !== undefined,
        assets: manifestAssets
    };

    if (bundle.original) {
        const originalName = `${manifest.name}.${bundle.format || 'bin'}`;
        manifest.original = originalName;
        zip.file(originalName, bundle.original, { binary: true });
    }
    if (bundle.document !== undefined) {
        zip.file('document.json', JSON.stringify(bundle.document, null, 2));
    }
    if (bundle.extra) manifest.extra = bundle.extra;

    zip.file(BUNDLE_MANIFEST, JSON.stringify(manifest, null, 2));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** 解包结果 */
export interface UnpackedDocument extends DocumentBundle {
    /** 资源清单（位置锚点完整保留） */
    manifest: BundleManifest;
}

/**
 * 解开一体包，还原文档结构与资源字节
 * @param input - bundle 字节
 * @returns 文档结构、资源与清单
 */
export async function unpackDocument(input: DocInput): Promise<UnpackedDocument> {
    const bytes = await toUint8Array(input);
    const zip = await JSZip.loadAsync(bytes);
    const manifestFile = zip.file(BUNDLE_MANIFEST);
    if (!manifestFile) {
        throw new TypeError('jdoc: 不是有效的 jdoc bundle（缺少 manifest.json）');
    }
    const manifest = JSON.parse(await manifestFile.async('string')) as BundleManifest;

    const assets: DocAsset[] = [];
    for (const entry of manifest.assets || []) {
        const file = zip.file(entry.path);
        if (!file) continue;
        const asset: DocAsset = {
            id: entry.id,
            name: entry.name,
            mime: entry.mime,
            data: new Uint8Array(await file.async('uint8array')),
            anchor: entry.anchor
        };
        if (entry.part) asset.part = entry.part;
        if (entry.width != null) asset.width = entry.width;
        if (entry.height != null) asset.height = entry.height;
        if (entry.alt) asset.alt = entry.alt;
        assets.push(asset);
    }

    const docFile = zip.file('document.json');
    const originalFile = manifest.original ? zip.file(manifest.original) : null;

    const result: UnpackedDocument = {
        format: manifest.format,
        name: manifest.name,
        assets,
        manifest
    };
    if (docFile) result.document = JSON.parse(await docFile.async('string'));
    if (originalFile) result.original = new Uint8Array(await originalFile.async('uint8array'));
    if (manifest.extra) result.extra = manifest.extra;
    return result;
}

/**
 * 判断字节是否为 jdoc bundle（避免对普通文档误判）
 * @param input - 文件字节
 */
export async function isDocumentBundle(input: DocInput): Promise<boolean> {
    try {
        const bytes = await toUint8Array(input);
        if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false; // 非 ZIP
        const zip = await JSZip.loadAsync(bytes);
        return !!zip.file(BUNDLE_MANIFEST);
    } catch {
        return false;
    }
}

/**
 * 把 bundle 落盘到目录（Node 专用）：文档、document.json、资源分文件存放
 *
 * 适合「另存为一个文件夹」的场景——资源各自独立成文件，便于人工查看与二次加工。
 *
 * @param bundle - 一体包内容
 * @param dir - 目标目录
 * @returns 写出的文件路径列表
 *
 * @example
 * ```js
 * await writeDocumentBundle('out/合同', { format:'docx', name:'合同', document, assets });
 * // => out/合同/合同.docx、out/合同/document.json、out/合同/assets/docx-0001.png
 * ```
 */
export async function writeDocumentBundle(bundle: DocumentBundle, dir: string): Promise<string[]> {
    const { mkdir, writeFile } = await import('node:fs/promises');
    const { join } = await import('node:path');

    await mkdir(join(dir, 'assets'), { recursive: true });
    const written: string[] = [];

    const manifestAssets: ManifestAsset[] = [];
    for (const [index, asset] of (bundle.assets || []).entries()) {
        const id = asset.id || createAssetId(bundle.format, index);
        const fileName = `${id}${extFor(asset.mime || mimeOf(asset.part || asset.name), asset.name)}`;
        const path = join(dir, 'assets', fileName);
        await writeFile(path, asset.data);
        written.push(path);
        const entry: ManifestAsset = {
            id,
            name: asset.name || fileName,
            mime: asset.mime || 'application/octet-stream',
            path: `assets/${fileName}`,
            anchor: asset.anchor,
            size: asset.data?.byteLength || 0
        };
        if (asset.part) entry.part = asset.part;
        manifestAssets.push(entry);
    }

    const manifest: BundleManifest = {
        kind: 'jdoc-bundle',
        version: BUNDLE_VERSION,
        format: bundle.format,
        name: bundle.name || 'document',
        hasDocument: bundle.document !== undefined,
        assets: manifestAssets
    };

    if (bundle.original) {
        const originalName = `${manifest.name}.${bundle.format || 'bin'}`;
        const path = join(dir, originalName);
        await writeFile(path, bundle.original);
        manifest.original = originalName;
        written.push(path);
    }
    if (bundle.document !== undefined) {
        const path = join(dir, 'document.json');
        await writeFile(path, JSON.stringify(bundle.document, null, 2), 'utf8');
        written.push(path);
    }
    const manifestPath = join(dir, BUNDLE_MANIFEST);
    await writeFile(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');
    written.push(manifestPath);
    return written;
}

/**
 * 列出 bundle 内每个资源的位置（不读字节，用于「资源 ↔ 文档位置」导航界面）
 * @param input - bundle 字节
 */
export async function listBundleAssets(input: DocInput): Promise<ManifestAsset[]> {
    const bytes = await toUint8Array(input);
    const zip = await JSZip.loadAsync(bytes);
    const manifestFile = zip.file(BUNDLE_MANIFEST);
    if (!manifestFile) throw new TypeError('jdoc: 不是有效的 jdoc bundle（缺少 manifest.json）');
    const manifest = JSON.parse(await manifestFile.async('string')) as BundleManifest;
    return manifest.assets || [];
}