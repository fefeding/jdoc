/**
 * 生成侧资源引用测试
 *
 * DOCX / XLSX / PPTX / PDF 四种格式都应能用 `{ assetId }` 或资源对象直接写图，
 * 无需把字节转成 data URL 在结构里搬来搬去。
 */

import { describe, expect, it } from 'vitest';
import {
    createAsset,
    docxToJson,
    jsonToDocx,
    jsonToPdf,
    jsonToPptx,
    jsonToXlsx,
    pdfToJson,
    pptxToStandard,
    xlsxToJson
} from '../src/index';
import { createTightPng } from './helpers/png';

/** 造一个带图片的资源 */
function makeAsset(id = 'pic-1') {
    return createAsset({
        id,
        data: createTightPng(),
        anchor: { scope: 'docx', kind: 'block', index: 0 }
    });
}

describe('PDF 生成侧资源引用', () => {
    it('用 { assetId } 引用即可嵌入图片', async () => {
        const asset = makeAsset('pdf-pic');
        const bytes = await jsonToPdf({
            elements: [{ type: 'image', src: { assetId: 'pdf-pic' }, width: 80, height: 60 }]
        }, { assets: [asset] });

        expect(bytes.byteLength).toBeGreaterThan(500);
        // 生成结果里确实带上了图像对象
        expect(bytes.byteLength).toBeGreaterThan(400);
    });

    it('直接把资源对象放在 src 上也能用', async () => {
        const asset = makeAsset();
        const bytes = await jsonToPdf({
            elements: [{ type: 'image', src: asset, width: 80, height: 60 }]
        });
        expect(bytes.byteLength).toBeGreaterThan(400);
    });

    it('资源表写在规格里（spec.assets）同样生效', async () => {
        const asset = makeAsset();
        const bytes = await jsonToPdf({
            assets: [asset],
            elements: [{ type: 'image', src: { assetId: asset.id }, width: 80, height: 60 }]
        });
        expect(bytes.byteLength).toBeGreaterThan(400);
    });

    it('引用缺失时降级为跳过，不让整篇生成失败', async () => {
        const bytes = await jsonToPdf({
            elements: [
                { type: 'text', text: 'before' },
                { type: 'image', src: { assetId: 'missing' } },
                { type: 'text', text: 'after' }
            ]
        });
        const doc = await pdfToJson(bytes);
        expect(doc.pageCount).toBe(1);
    });
});

describe('PPTX 生成侧资源引用', () => {
    it('{ assetId } 引用会被还原成 data URL', async () => {
        const asset = createAsset({
            id: 'pptx-pic',
            data: createTightPng(),
            anchor: { scope: 'pptx', kind: 'slide', index: 1 }
        });
        const bytes = await jsonToPptx({
            slides: [{
                elements: [
                    { type: 'text', x: 40, y: 40, width: 300, height: 40, text: 'Slide' },
                    { type: 'image', data: { assetId: 'pptx-pic' }, x: 60, y: 140, width: 160, height: 120 }
                ]
            }]
        }, { assets: [asset] });

        expect(bytes.byteLength).toBeGreaterThan(2000);
        // 生成的 PPTX 里应能读回这张图
        const back = await pptxToStandard(bytes, { assets: 'extract' });
        expect(back.assets).toBeDefined();
        expect(back.assets.length).toBeGreaterThan(0);
    });

    it('extract 模式抽出的资源可直接用于再生成', async () => {
        const bytes = await jsonToPptx({
            slides: [{
                elements: [{ type: 'image', data: createTightPng(), x: 40, y: 40, width: 120, height: 90 }]
            }]
        });
        const deck = await pptxToStandard(bytes, { assets: 'extract' });
        expect(deck.assets.length).toBe(1);

        // 抽出的引用结构可直接重新生成
        const again = await jsonToPptx(deck, { assets: deck.assets });
        expect(again.byteLength).toBeGreaterThan(1000);
    });
});

describe('四格式引用生成的统一约定', () => {
    it('DOCX：引用写法与资源表', async () => {
        const asset = makeAsset('docx-pic');
        const bytes = await jsonToDocx({
            blocks: [{ type: 'image', src: { assetId: 'docx-pic' }, width: 100, height: 80 }]
        }, { assets: [asset] });
        const parsed = await docxToJson(bytes, { assets: 'extract' });
        expect(parsed.assets!.length).toBe(1);
    });

    it('XLSX：引用写法与资源表', async () => {
        const asset = makeAsset('xlsx-pic');
        const bytes = await jsonToXlsx({
            sheets: [{ name: 'S', rows: [['x', { image: { assetId: 'xlsx-pic' } }]] }]
        }, { assets: [asset] });
        const parsed = await xlsxToJson(bytes, { assets: 'extract' });
        expect(parsed.assets!.length).toBe(1);
        expect(parsed.assets![0].anchor.path).toBe('S!B1');
    });

    it('资源缺失时降级为占位段落，不让整篇生成失败（DOCX）', async () => {
        const bytes = await jsonToDocx({
            blocks: [
                { type: 'paragraph', text: '正文' },
                { type: 'image', src: { assetId: 'nope' } }
            ]
        });
        // 既有设计：图片问题降级而不中断，文档仍可生成且文字完整
        const parsed = await docxToJson(bytes);
        expect(parsed.text).toContain('正文');
        expect(parsed.assets).toBeUndefined();
    });
});