import { describe, expect, it } from 'vitest';
import { detectFormat } from '../src/index';
import { jsonToPptx, pptxToFiles, pptxToHtml, pptxToJson, pptxToStandard } from '../src/pptx/index';

describe('pptx：生成', () => {
    it('jsonToPptx 返回 Uint8Array 且是合法 ZIP', async () => {
        const bytes = await jsonToPptx({
            slides: [{ elements: [{ type: 'text', text: 'Hello PPTX', x: 50, y: 50, width: 300, height: 40 }] }]
        });
        expect(bytes).toBeInstanceOf(Uint8Array);
        expect(bytes.byteLength).toBeGreaterThan(1000);
        expect(bytes[0]).toBe(0x50);
        expect(bytes[1]).toBe(0x4b);
    });

    it('支持 PPTXComposer 链式构建', async () => {
        const { PPTXComposer } = await import('../src/pptx/index');
        const composer = new PPTXComposer();
        composer.title('Composed').addSlide((slide: Record<string, (fn: unknown) => unknown>) => {
            slide.addText({ text: 'Composed slide', x: 20, y: 20, width: 200, height: 30 });
        });
        const bytes = (await composer.save()) as Uint8Array;
        expect(bytes.byteLength).toBeGreaterThan(1000);
    });
});

describe('pptx：解析', () => {
    it('pptxToStandard 读回文档结构', async () => {
        const bytes = await jsonToPptx({
            slides: [{ elements: [{ type: 'text', text: 'Round Trip', x: 10, y: 10, width: 200, height: 30 }] }]
        });
        const doc = (await pptxToStandard(bytes)) as Record<string, unknown>;
        expect(doc).toBeTruthy();
        expect(doc.slideSize).toBeTruthy();
        expect(Array.isArray(doc.slides)).toBe(true);
        expect((doc.slides as unknown[]).length).toBeGreaterThan(0);
        expect(JSON.stringify(doc)).toContain('Round Trip');
    });

    it('pptxToHtml 返回幻灯片数组', async () => {
        const bytes = await jsonToPptx({ slides: [{ elements: [] }] });
        const result = await pptxToHtml(bytes);
        expect(Array.isArray(result.slides)).toBe(true);
        expect(result.slides.length).toBeGreaterThan(0);
    });

    it('pptxToJson 返回结构化结果', async () => {
        const bytes = await jsonToPptx({ slides: [{ elements: [] }] });
        const result = await pptxToJson(bytes);
        expect(Array.isArray(result.slides)).toBe(true);
    });

    it('pptxToFiles 列出内部部件', async () => {
        const bytes = await jsonToPptx({ slides: [{ elements: [] }] });
        const result = await pptxToFiles(bytes);
        expect(result.files.some((f: { name: string }) => f.name.includes('ppt/'))).toBe(true);
    });

    it('支持 base64 字符串输入', async () => {
        const bytes = await jsonToPptx({ slides: [{ elements: [] }] });
        const base64 = Buffer.from(bytes).toString('base64');
        const doc = await pptxToStandard(base64);
        expect(doc).toBeTruthy();
    });
});

describe('detectFormat：真实文件识别', () => {
    it('识别生成的四种文件', async () => {
        const pptxBytes = await jsonToPptx({ slides: [{ elements: [] }] });
        expect(await detectFormat(pptxBytes)).toBe('pptx');

        const { jsonToDocx } = await import('../src/docx/index');
        const docxBytes = await jsonToDocx([{ type: 'paragraph', text: 'hi' }]);
        expect(await detectFormat(docxBytes)).toBe('docx');

        const { jsonToXlsx } = await import('../src/xlsx/index');
        const xlsxBytes = await jsonToXlsx([{ name: 'S', rows: [['a']] }]);
        expect(await detectFormat(xlsxBytes)).toBe('xlsx');

        const { jsonToPdf } = await import('../src/pdf/index');
        const pdfBytes = await jsonToPdf({ elements: [{ type: 'text', text: 'hi' }] });
        expect(await detectFormat(pdfBytes)).toBe('pdf');
    });
});
