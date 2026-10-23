import { describe, expect, it } from 'vitest';
import { jsonToPdf } from '../src/pdf/index';
import { pdfToBlocks } from '../src/convert/pdf';
import { convert } from '../src/convert/index';
import { encodePng } from '../src/core/png';
import type { DocxBlock } from '../src/docx/types';

/** 造一张 32x24 的纯红 PNG（不依赖任何外部资源；尺寸需过提取的最小像素阈值） */
async function redPng(): Promise<Uint8Array> {
    const width = 32;
    const height = 24;
    const rgb = new Uint8Array(width * height * 3);
    for (let i = 0; i < width * height; i++) {
        rgb[i * 3] = 220;
        rgb[i * 3 + 1] = 40;
        rgb[i * 3 + 2] = 40;
    }
    return encodePng({ data: rgb, width, height, kind: 2 });
}

/** 生成「文字 / 图 / 文字」的 PDF */
async function pdfWithImage(): Promise<Uint8Array> {
    const png = await redPng();
    return jsonToPdf({
        elements: [
            { type: 'heading', level: 1, text: 'Image Section' },
            { type: 'text', text: 'Text before the image.' },
            { type: 'image', src: png, width: 240, height: 160 },
            { type: 'text', text: 'Text after the image.' }
        ]
    });
}

describe('pdf：页面图像还原', () => {
    it('PNG 编码器产出可被系统识别的 PNG', async () => {
        const png = await redPng();
        expect(png[0]).toBe(0x89);
        expect(png[1]).toBe(0x50);
        expect(png.byteLength).toBeGreaterThan(50);
    });

    it('pdf -> blocks 还原出图片块与资源表', async () => {
        const pdf = await pdfWithImage();
        const { blocks, assets } = await pdfToBlocks(pdf, { quality: 'fast' });
        const images = blocks.filter((b) => b.type === 'image');
        expect(images.length).toBe(1);
        expect(assets?.length).toBe(1);
        expect(assets![0].mime).toBe('image/png');
        // 显示尺寸沿用 PDF 里的绘制尺寸（磅）
        expect(Math.round((images[0] as { width?: number }).width || 0)).toBe(240);
    });

    it('图片按阅读顺序插在两段文字之间', async () => {
        const pdf = await pdfWithImage();
        const { blocks } = await pdfToBlocks(pdf, { quality: 'fast' });
        const kinds = blocks.map((b) => b.type);
        const imgIndex = kinds.indexOf('image');
        expect(imgIndex).toBeGreaterThan(0);
        const before = blocks.slice(0, imgIndex).some((b) => b.type === 'paragraph');
        const after = blocks.slice(imgIndex + 1).some((b) => b.type === 'paragraph');
        expect(before).toBe(true);
        expect(after).toBe(true);
    });

    it('pdf -> md / html 内联出图片', async () => {
        const pdf = await pdfWithImage();
        const md = (await convert(pdf, 'md', { from: 'pdf' })) as string;
        expect(md).toContain('data:image/png;base64,');

        const html = (await convert(pdf, 'html', { from: 'pdf' })) as string;
        expect(html).toContain('<img');
        expect(html).toContain('data:image/png;base64,');
    });

    it('images:false 时不产图片块', async () => {
        const pdf = await pdfWithImage();
        const { blocks } = await pdfToBlocks(pdf, { quality: 'fast', images: false });
        expect(blocks.some((b: DocxBlock) => b.type === 'image')).toBe(false);
    });
});
