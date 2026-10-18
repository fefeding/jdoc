/**
 * XLSX / PPTX / PDF 的二进制资源接入测试
 *
 * 覆盖：单元格图片的提取与锚点、引用式生成往返、
 *       PPTX 元素图片抽取、PDF 附件与页面图像提取
 */

import { describe, expect, it } from 'vitest';
import {
    assetToDataUrl,
    jsonToXlsx,
    packDocument,
    pdfToJson,
    pptxToStandard,
    unpackDocument,
    xlsxToJson,
    type DocAsset
} from '../src/index';
import { createTightPng } from './helpers/png';

describe('XLSX 单元格图片', () => {
    /** 造一个带图片的工作簿：图片锚定在 B2 */
    async function buildSheetWithImage(): Promise<Uint8Array> {
        return jsonToXlsx({
            sheets: [
                {
                    name: '封面',
                    rows: [
                        ['标题', { image: createTightPng(), imageWidth: 120, imageHeight: 90 }]
                    ]
                }
            ]
        });
    }

    it('默认不提取图片，保持轻量解析', async () => {
        const file = await buildSheetWithImage();
        const parsed = await xlsxToJson(file);
        expect(parsed.assets).toBeUndefined();
    });

    it('assets: extract 时图片进资源表并锚定到单元格', async () => {
        const file = await buildSheetWithImage();
        const parsed = await xlsxToJson(file, { assets: 'extract' });

        expect(parsed.assets).toBeDefined();
        expect(parsed.assets!.length).toBe(1);

        const asset = parsed.assets![0];
        expect(asset.id).toMatch(/^xlsx-\d{4}$/);
        expect(asset.mime).toBe('image/png');
        expect(asset.data.byteLength).toBeGreaterThan(0);
        // 锚点：哪张表的哪个单元格
        expect(asset.anchor.scope).toBe('xlsx');
        expect(asset.anchor.kind).toBe('cell');
        expect(asset.anchor.path).toMatch(/^封面!B1$/);

        // 单元格上是引用，不是 data URL
        const cell = parsed.sheets[0].rows[0] as { cells: { image?: unknown; assetId?: string }[] };
        expect(cell.cells[1].assetId).toBe(asset.id);
        expect(typeof cell.cells[1].image).toBe('object');
    });

    it('assets: both 时单元格带 data URL 且资源表齐全', async () => {
        const file = await buildSheetWithImage();
        const parsed = await xlsxToJson(file, { assets: 'both' });
        const cell = parsed.sheets[0].rows[0] as { cells: { image?: unknown }[] };

        expect(String(cell.cells[1].image)).toMatch(/^data:image\/png;base64,/);
        expect(parsed.assets!.length).toBe(1);
    });

    it('引用式往返：解析 → 生成，图片不丢', async () => {
        const file = await buildSheetWithImage();
        const parsed = await xlsxToJson(file, { assets: 'extract' });

        const rebuilt = await jsonToXlsx(
            { sheets: parsed.sheets, assets: parsed.assets },
            { assets: parsed.assets }
        );
        const reparsed = await xlsxToJson(rebuilt, { assets: 'extract' });

        expect(reparsed.assets!.length).toBe(1);
        expect(reparsed.assets![0].data.byteLength).toBe(parsed.assets![0].data.byteLength);
    });

    it('HTML 预览能显示图片单元格', async () => {
        const file = await buildSheetWithImage();
        const parsed = await xlsxToJson(file, { assets: 'extract' });
        const { xlsxToHtml } = await import('../src/index');
        const { html } = await xlsxToHtml(file, { assets: 'extract' });
        expect(html).toContain('<img');
        expect(html).toContain('data:image/png;base64,');
        expect(parsed.assets!.length).toBe(1);
    });
});

describe('PPTX 图片资源', () => {
    it('抽取元素图片为资源，并记录所属幻灯片', async () => {
        // 直接构造标准结构：ppt-parser 会把图片放在元素 data 上
        const png = createTightPng();
        const doc = {
            slides: [
                { elements: [{ type: 'text', text: '标题' }, { type: 'image', data: assetToDataUrl(png, 'image/png') }] },
                { elements: [{ type: 'image', data: assetToDataUrl(png, 'image/png') }] }
            ]
        };
        const { extractPptxAssets } = await import('../src/pptx/assets');
        const assets: DocAsset[] = extractPptxAssets(doc, 'extract');

        expect(assets.length).toBe(2);
        expect(assets[0].id).toBe('pptx-0001');
        expect(assets[0].anchor.scope).toBe('pptx');
        expect(assets[0].anchor.kind).toBe('slide');
        expect(assets[0].anchor.index).toBe(1);
        expect(assets[1].anchor.index).toBe(2);
        expect(assets[0].data.byteLength).toBe(png.byteLength);

        // extract 模式清掉元素上的 data，避免 JSON 膨胀
        const first = doc.slides[0].elements[1] as { data?: unknown; assetId?: string };
        expect(first.data).toBeUndefined();
        expect(first.assetId).toBe('pptx-0001');
    });

    it('both 模式保留元素 data 以便直接显示', async () => {
        const png = createTightPng();
        const doc = {
            slides: [{ elements: [{ type: 'image', data: assetToDataUrl(png, 'image/png') }] }]
        };
        const { extractPptxAssets } = await import('../src/pptx/assets');
        const assets = extractPptxAssets(doc, 'both');

        expect(assets.length).toBe(1);
        const element = doc.slides[0].elements[0] as { data?: unknown; assetId?: string };
        expect(String(element.data)).toMatch(/^data:image\/png;base64,/);
        expect(element.assetId).toBe('pptx-0001');
    });

    it('pptxToStandard 透传 assets 选项而不影响原有返回', async () => {
        const { jsonToPptx } = await import('../src/index');
        const bytes = await jsonToPptx({
            slides: [{ elements: [{ type: 'text', x: 40, y: 40, width: 300, height: 40, text: 'Hello' }] }]
        });
        const doc = await pptxToStandard(bytes);
        expect(doc.slides.length).toBeGreaterThan(0);
        // 无图片时不产生资源表
        expect(doc.assets).toBeUndefined();
    });
});

describe('PDF 内嵌资源', () => {
    /**
     * 造一个带内嵌附件的 PDF
     *
     * 注意：pdf-lib 的 `attach()` 写的是页面级 AF 数组，pdfjs 读的是
     * catalog → Names → EmbeddedFiles 名字树，两者不通用，这里手工按规范构造。
     */
    async function buildPdfWithAttachment(): Promise<Uint8Array> {
        const { PDFDocument, PDFName, PDFString } = await import('pdf-lib');
        const pdf = await PDFDocument.create();
        pdf.addPage();
        const ctx = pdf.context;
        const stream = ctx.stream(createTightPng(), { Type: 'EmbeddedFile' });
        const spec = ctx.obj({
            Type: '/Filespec',
            F: PDFString.of('logo.png'),
            UF: PDFString.of('logo.png'),
            EF: ctx.obj({ F: stream, UF: stream })
        });
        const names = ctx.obj({});
        names.set(
            PDFName.of('EmbeddedFiles'),
            ctx.obj({ Names: ctx.obj([PDFString.of('logo.png'), spec]) })
        );
        pdf.catalog.set(PDFName.of('Names'), names);
        return new Uint8Array(await pdf.save());
    }

    it('默认不提取内嵌资源', async () => {
        const file = await buildPdfWithAttachment();
        const doc = await pdfToJson(file);
        expect(doc.assets).toBeUndefined();
    });

    it('assets: extract 时附件进资源表并带名称与位置', async () => {
        const file = await buildPdfWithAttachment();
        const doc = await pdfToJson(file, { assets: 'extract' });

        expect(doc.assets).toBeDefined();
        expect(doc.assets!.length).toBeGreaterThan(0);

        const asset = doc.assets![0];
        expect(asset.id).toMatch(/^pdf-\d{4}$/);
        expect(asset.name).toBe('logo.png');
        expect(asset.anchor.scope).toBe('pdf');
        expect(asset.anchor.kind).toBe('other');
        expect(asset.anchor.label).toContain('logo.png');
        expect(asset.data.byteLength).toBeGreaterThan(0);
    });

    it('忽略附件不影响正文解析', async () => {
        const file = await buildPdfWithAttachment();
        const doc = await pdfToJson(file, { assets: 'extract' });
        expect(doc.pageCount).toBe(1);
        expect(doc.pages?.length).toBe(1);
    });

    it('页面图像提取：扫描件页内的插图可被抽出并锚定页码', async () => {
        const { extractPdfAssets } = await import('../src/pdf/assets');
        const png = createTightPng();
        // 构造一个最小 pdfjs 文档替身：一个附件 + 两页，每页画一张图
        const fake = {
            numPages: 2,
            getAttachments: async () => ({
                a1: { filename: 'note.txt', content: new TextEncoder().encode('hello') }
            }),
            getPage: async (pageNo: number) => ({
                getOperatorList: async () => ({ fnArray: [85], argsArray: [['Im1']] }),
                objs: {
                    get: (key: string) =>
                        key === 'Im1' ? { data: png.buffer, width: 1, height: 1 } : undefined
                },
                commonObjs: { get: () => undefined }
            })
        };

        const assets = await extractPdfAssets(fake as never, { mode: 'extract' });
        // 1 个附件 + 2 页各 1 张图
        expect(assets.length).toBe(3);

        const attachment = assets.find((a) => a.anchor.kind === 'other')!;
        expect(attachment.name).toBe('note.txt');
        expect(attachment.mime).toBe('text/plain');

        const pageImages = assets.filter((a) => a.anchor.kind === 'page');
        expect(pageImages.length).toBe(2);
        expect(pageImages[0].anchor.index).toBe(1);
        expect(pageImages[1].anchor.index).toBe(2);
        expect(pageImages[0].anchor.label).toBe('第 1 页');
        expect(pageImages[0].data.byteLength).toBe(png.byteLength);
    });
});

describe('跨格式统一：bundle 打包', () => {
    it('XLSX 工作簿也能整体打包并还原', async () => {
        const file = await jsonToXlsx({
            sheets: [{ name: '数据', rows: [['列', { image: createTightPng() }]] }]
        });
        const parsed = await xlsxToJson(file, { assets: 'extract' });

        const zip = await packDocument({
            format: 'xlsx',
            name: '数据',
            document: parsed,
            assets: parsed.assets,
            original: file
        });
        const restored = await unpackDocument(zip);

        expect(restored.format).toBe('xlsx');
        expect(restored.assets.length).toBe(1);
        expect(restored.assets[0].anchor.kind).toBe('cell');
        // 还原后能直接重新生成
        const again = await jsonToXlsx(restored.document as never, { assets: restored.assets });
        const reparsed = await xlsxToJson(again, { assets: 'extract' });
        expect(reparsed.assets!.length).toBe(1);
    });
});