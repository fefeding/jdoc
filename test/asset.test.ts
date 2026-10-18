/**
 * 二进制资源（图片等）测试
 *
 * 覆盖三条主线：
 *   1. 解析：资源能被抽出成资源表，块上留引用，并带位置锚点
 *   2. 取用：同一份字节能按端拿到 data URL / Blob / 字节 / 流
 *   3. 一体保存：文档 + 资源能打成 bundle 并原样解回，往返生成不缺图
 */

import { describe, expect, it } from 'vitest';
import {
    assetToArrayBuffer,
    assetToBase64,
    assetToBlob,
    assetToDataUrl,
    assetToNodeStream,
    createAsset,
    createAssetId,
    docxToJson,
    getAsset,
    isDocumentBundle,
    jsonToDocx,
    listBundleAssets,
    mimeOf,
    packDocument,
    resolveAssetBytes,
    sniffMime,
    summarizeAssets,
    unpackDocument
} from '../src/index';
import { createTightPng } from './helpers/png';

/** 构造一个带图片的 docx（用生成器再解析，保证往返链路一致） */
async function buildDocxWithImage(): Promise<Uint8Array> {
    const png = createTightPng();
    return jsonToDocx({
        metadata: { title: '带图文档' },
        blocks: [
            { type: 'heading', level: 1, text: '封面' },
            { type: 'image', src: png, width: 120, height: 80, alt: '示意图' },
            { type: 'paragraph', text: '正文说明' }
        ]
    });
}

describe('二进制资源：解析提取', () => {
    it('assets: extract 时字节进资源表，块上留 assetId 引用', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file, { assets: 'extract' });

        expect(parsed.assets).toBeDefined();
        expect(parsed.assets!.length).toBe(1);

        const asset = parsed.assets![0];
        expect(asset.id).toMatch(/^docx-\d{4}$/);
        expect(asset.mime).toBe('image/png');
        expect(asset.data).toBeInstanceOf(Uint8Array);
        expect(asset.data.byteLength).toBeGreaterThan(0);
        expect(asset.part).toMatch(/^word\/media\//);

        // 块上是引用而非 data URL，JSON 里不再塞 base64
        const image = parsed.blocks.find((b) => b.type === 'image');
        expect(image).toBeDefined();
        expect(image!.type === 'image' && image!.assetId).toBe(asset.id);
        expect(JSON.stringify(parsed.blocks)).not.toContain('data:image');
    });

    it('资源带位置锚点，能反查它在文档中的落点', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file, { assets: 'extract' });
        const asset = parsed.assets![0];

        expect(asset.anchor.scope).toBe('docx');
        expect(asset.anchor.kind).toBe('block');
        // 第 2 个块（heading 是第 1 个，图片紧跟其后）
        expect(asset.anchor.index).toBe(1);
        expect(asset.anchor.path).toBe('blocks[1]');
        expect(asset.anchor.label).toContain('第 2');
    });

    it('assets: inline（默认）保持原有行为，块上是 data URL', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file);
        const image = parsed.blocks.find((b) => b.type === 'image');

        expect(image!.type === 'image' && String(image!.src).startsWith('data:image/png;base64,')).toBe(true);
        expect(parsed.assets).toBeUndefined();
    });

    it('assets: both 时两边都给——前端可显示，后台可取字节', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file, { assets: 'both' });
        const image = parsed.blocks.find((b) => b.type === 'image');

        expect(parsed.assets!.length).toBe(1);
        expect(String(image!.type === 'image' && image!.src)).toMatch(/^data:image\/png;base64,/);
        expect(image!.type === 'image' && image!.assetId).toBe(parsed.assets![0].id);
    });

    it('同一张图多处引用只存一份字节', async () => {
        const png = createTightPng();
        const file = await jsonToDocx({
            blocks: [
                { type: 'image', src: png },
                { type: 'paragraph', text: '中间' },
                { type: 'image', src: png }
            ]
        });
        const parsed = await docxToJson(file, { assets: 'extract' });

        const images = parsed.blocks.filter((b) => b.type === 'image');
        expect(images.length).toBe(2);
        // 字节只存一份，两个块引用同一资源
        expect(parsed.assets!.length).toBe(1);
        expect(images[0].type === 'image' && images[1].type === 'image' &&
            images[0].assetId === images[1].assetId).toBe(true);
    });
});

describe('二进制资源：多端取用', () => {
    const asset = createAsset({
        id: createAssetId('test', 1),
        data: createTightPng(),
        anchor: { scope: 'docx', kind: 'block', index: 0 }
    });

    it('浏览器形态：data URL 与 Blob', () => {
        expect(assetToDataUrl(asset)).toMatch(/^data:image\/png;base64,/);
        const blob = assetToBlob(asset);
        expect(blob.type).toBe('image/png');
        expect(blob.size).toBe(asset.data.byteLength);
    });

    it('后台形态：字节、base64、ArrayBuffer', () => {
        expect(assetToBase64(asset)).toMatch(/^[A-Za-z0-9+/]+=*$/);
        expect(assetToArrayBuffer(asset).byteLength).toBe(asset.data.byteLength);
    });

    it('后台形态：Node 可读流可直接拿到字节', async () => {
        const stream = assetToNodeStream(asset) as { readable: boolean };
        expect(stream).toBeTruthy();
        // Readable / ReadableStream 都提供读取能力
        const hasReader = typeof (stream as { getReader?: unknown }).getReader === 'function';
        const hasPipe = typeof (stream as { pipe?: unknown }).pipe === 'function';
        expect(hasReader || hasPipe).toBe(true);
    });

    it('资源是纯数据，可直接 JSON 序列化（去掉字节后落库）', () => {
        const { data, ...meta } = asset;
        expect(data).toBeInstanceOf(Uint8Array);
        const json = JSON.stringify(meta);
        expect(JSON.parse(json).anchor.kind).toBe('block');
    });

    it('按魔数识别 MIME，无需用户标注', () => {
        expect(sniffMime(createTightPng())).toBe('image/png');
        expect(sniffMime(new Uint8Array([0xff, 0xd8, 0xff, 0xe0]))).toBe('image/jpeg');
        expect(sniffMime(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]))).toBe('application/pdf');
        expect(mimeOf('word/media/image3.JPG')).toBe('image/jpeg');
    });

    it('资源表工具：按 id 取用与体积汇总', () => {
        expect(getAsset([asset], asset.id)).toBe(asset);
        expect(getAsset({ assets: [asset] }, asset.id)).toBe(asset);
        expect(getAsset([asset], 'missing')).toBeUndefined();
        expect(summarizeAssets([asset])).toEqual({ count: 1, bytes: asset.data.byteLength });
    });
});

describe('二进制资源：生成侧引用', () => {
    it('块上用 { assetId } 引用即可生成，不必内联 base64', async () => {
        const png = createTightPng();
        const asset = createAsset({
            id: 'logo-1',
            data: png,
            anchor: { scope: 'docx', kind: 'block', index: 0 }
        });

        const bytes = await jsonToDocx({
            blocks: [{ type: 'image', src: { assetId: 'logo-1' }, width: 100, height: 60 }],
            assets: [asset]
        });
        expect(bytes.byteLength).toBeGreaterThan(1000);

        // 生成结果里确实含有图片
        const parsed = await docxToJson(bytes, { assets: 'extract' });
        expect(parsed.assets!.length).toBe(1);
        expect(parsed.assets![0].mime).toBe('image/png');
    });

    it('resolveAssetBytes 统一解析各种写法', async () => {
        const png = createTightPng();
        const asset = createAsset({
            id: 'a1',
            data: png,
            anchor: { scope: 'docx', kind: 'block', index: 0 }
        });

        const byRef = await resolveAssetBytes({ assetId: 'a1' }, [asset]);
        const byObject = await resolveAssetBytes(asset);
        const byDataUrl = await resolveAssetBytes(assetToDataUrl(asset));
        const byBytes = await resolveAssetBytes(png);

        expect(byRef!.bytes.byteLength).toBe(png.byteLength);
        expect(byObject!.bytes.byteLength).toBe(png.byteLength);
        expect(byDataUrl!.mime).toBe('image/png');
        expect(byBytes!.mime).toBe('image/png'); // 靠魔数识别
        expect(await resolveAssetBytes({ assetId: 'nope' }, [asset])).toBeNull();
    });

    it('解析 → 改 → 生成往返，图片不丢', async () => {
        const original = await buildDocxWithImage();
        const parsed = await docxToJson(original, { assets: 'extract' });

        // 改一处文字，图片块保持引用
        // 解析出的段落以 runs 承载文字，改动也要落在 runs 上
        const blocks = parsed.blocks.map((block) =>
            block.type === 'paragraph'
                ? { ...block, text: undefined, runs: [{ text: '改过的正文' }] }
                : block
        );
        const rebuilt = await jsonToDocx({ ...parsed, blocks });

        const reparsed = await docxToJson(rebuilt, { assets: 'extract' });
        expect(reparsed.assets!.length).toBe(1);
        expect(reparsed.text).toContain('改过的正文');
        // 图片字节与原件一致（尺寸元数据允许有偏差，比对字节）
        expect(reparsed.assets![0].data.byteLength).toBe(parsed.assets![0].data.byteLength);
    });
});

describe('文档一体保存（bundle）', () => {
    it('文档 + 资源打成 bundle，解包后结构与字节一致', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file, { assets: 'extract' });

        const zip = await packDocument({
            format: 'docx',
            name: '带图文档',
            document: parsed,
            assets: parsed.assets,
            original: file
        });
        expect(zip.byteLength).toBeGreaterThan(0);
        expect(await isDocumentBundle(zip)).toBe(true);

        const restored = await unpackDocument(zip);
        expect(restored.format).toBe('docx');
        expect(restored.name).toBe('带图文档');
        expect(restored.assets.length).toBe(1);
        expect(restored.assets[0].data.byteLength).toBe(parsed.assets![0].data.byteLength);
        // 位置锚点在打包/解包后仍在
        expect(restored.assets[0].anchor.kind).toBe('block');
        expect(restored.assets[0].anchor.index).toBe(1);
        // 原始文件也一并带走
        expect(restored.original!.byteLength).toBe(file.byteLength);
        // 文档结构可回读
        expect((restored.document as { blocks: unknown[] }).blocks.length).toBe(parsed.blocks.length);
    });

    it('listBundleAssets 可列出资源及其在文档中的位置', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file, { assets: 'extract' });
        const zip = await packDocument({
            format: 'docx',
            name: '文档',
            document: parsed,
            assets: parsed.assets
        });

        const list = await listBundleAssets(zip);
        expect(list.length).toBe(1);
        expect(list[0].path).toMatch(/^assets\/docx-\d{4}\.png$/);
        expect(list[0].anchor.index).toBe(1);
        expect(list[0].size).toBeGreaterThan(0);
    });

    it('从 bundle 还原的文档可直接再生成 DOCX', async () => {
        const file = await buildDocxWithImage();
        const parsed = await docxToJson(file, { assets: 'extract' });
        const zip = await packDocument({
            format: 'docx',
            name: '可还原',
            document: parsed,
            assets: parsed.assets
        });

        const restored = await unpackDocument(zip);
        const bytes = await jsonToDocx(restored.document as never, { assets: restored.assets });
        const reparsed = await docxToJson(bytes, { assets: 'extract' });
        expect(reparsed.assets!.length).toBe(1);
    });

    it('非 bundle 文件不会被误判', async () => {
        const file = await buildDocxWithImage();
        expect(await isDocumentBundle(file)).toBe(false);
    });
});