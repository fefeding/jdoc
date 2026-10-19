import { describe, expect, it } from 'vitest';
import { jsonToPdf, editPdf, mergePdfs, splitPdf, pdfToText, pdfInfo } from '../src/pdf/index';

/** 造一个 3 页的源 PDF（英文，避免依赖中文字体） */
async function makeSource(): Promise<Uint8Array> {
    return jsonToPdf({
        pages: [
            { size: { width: 400, height: 600 }, elements: [{ type: 'text', x: 40, y: 50, size: 20, bold: true, text: 'Page One' }] },
            { size: { width: 400, height: 600 }, elements: [{ type: 'text', x: 40, y: 50, size: 20, bold: true, text: 'Page Two' }] },
            { size: { width: 400, height: 600 }, elements: [{ type: 'text', x: 40, y: 50, size: 20, bold: true, text: 'Page Three' }] }
        ]
    });
}

/** 最小合法 PNG（1x1 透明），用于图片叠加测试 */
const PNG_1X1 = Uint8Array.from(
    atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='),
    (c) => c.charCodeAt(0)
);

describe('pdf：叠加编辑', () => {
    it('矩形叠加不改变页数且保留原文', async () => {
        const src = await makeSource();
        const out = await editPdf(src, {
            overlays: [[{ type: 'rect', x: 40, y: 90, width: 200, height: 16, fill: '#ffe08a', opacity: 0.6 }]]
        });
        expect((await pdfInfo(out)).pageCount).toBe(3);
        const text = await pdfToText(out);
        expect(text).toContain('Page One');
        expect(text).toContain('Page Three');
    });

    it('文本叠加写入新内容，原内容仍在', async () => {
        const src = await makeSource();
        const out = await editPdf(src, {
            overlays: [[{ type: 'text', x: 40, y: 500, width: 300, size: 14, text: 'Signed by Alice' }]]
        });
        const text = await pdfToText(out);
        expect(text).toContain('Page One');
        expect(text).toContain('Signed by Alice');
    });

    it('同页多对象按数组顺序叠放', async () => {
        const src = await makeSource();
        const out = await editPdf(src, {
            overlays: [
                [
                    { type: 'rect', x: 10, y: 10, width: 50, height: 50, fill: '#ff0000' },
                    { type: 'rect', x: 20, y: 20, width: 50, height: 50, fill: '#00ff00' },
                    { type: 'image', x: 100, y: 100, width: 40, height: 40, data: PNG_1X1 }
                ]
            ]
        });
        expect((await pdfInfo(out)).pageCount).toBe(3);
    });

    it('图片格式声明为 jpg 但字节是 png 时报 pdf 错误', async () => {
        const src = await makeSource();
        await expect(
            editPdf(src, {
                overlays: [[{ type: 'image', x: 10, y: 10, width: 20, height: 20, data: PNG_1X1, format: 'jpg' }]]
            })
        ).rejects.toThrow(/pdf/);
    });

    it('图片旋转 90/180/270 度均保真', async () => {
        const src = await makeSource();
        for (const rotate of [90, 180, 270] as const) {
            const out = await editPdf(src, {
                overlays: [[{ type: 'image', x: 10, y: 10, width: 20, height: 40, data: PNG_1X1, rotate }]]
            });
            expect((await pdfInfo(out)).pageCount).toBe(3);
            expect(await pdfToText(out)).toContain('Page One');
        }
    });

    it('手绘路径叠加可保存（绝对坐标 path）', async () => {
        const src = await makeSource();
        const out = await editPdf(src, {
            overlays: [[{ type: 'path', x: 60, y: 200, d: 'M0 0 L40 20 L80 0', stroke: '#c00', strokeWidth: 2 }]]
        });
        expect((await pdfInfo(out)).pageCount).toBe(3);
    });

    it('路径包围盒跨多段命令仍正确（不会只算第一段）', async () => {
        const src = await makeSource();
        const d = 'M10 10 L40 10 L40 40 L10 40 Z M100 100 L200 100 L200 200';
        const out = await editPdf(src, { overlays: [[{ type: 'path', x: 20, y: 300, d, scale: 0.5 }]] });
        expect((await pdfInfo(out)).pageCount).toBe(3);
    });

    it('旋转页面：原文仍在（/Rotate 不影响内容流）', async () => {
        const src = await makeSource();
        const out = await editPdf(src, { rotate: [{ page: 0, degrees: 90 }, { page: 1, degrees: 180 }] });
        const text = await pdfToText(out);
        expect(text).toContain('Page One');
        expect(text).toContain('Page Two');
        expect(text).toContain('Page Three');
    });

    it('删除指定页', async () => {
        const src = await makeSource();
        const out = await editPdf(src, { removePages: [1] });
        expect((await pdfInfo(out)).pageCount).toBe(2);
        const text = await pdfToText(out);
        expect(text).toContain('Page One');
        expect(text).not.toContain('Page Two');
        expect(text).toContain('Page Three');
    });

    it('重排页序', async () => {
        const src = await makeSource();
        const out = await editPdf(src, { pageOrder: [2, 0, 1] });
        const text = await pdfToText(out);
        expect((await pdfInfo(out)).pageCount).toBe(3);
        expect(text.indexOf('Page Three')).toBeLessThan(text.indexOf('Page One'));
    });

    it('删页 + 重排组合：先删后重排', async () => {
        const src = await makeSource();
        const out = await editPdf(src, { removePages: [1], pageOrder: [1, 0] });
        const text = await pdfToText(out);
        expect((await pdfInfo(out)).pageCount).toBe(2);
        expect(text.indexOf('Page Three')).toBeLessThan(text.indexOf('Page One'));
    });

    it('元数据覆盖', async () => {
        const src = await makeSource();
        const out = await editPdf(src, { metadata: { title: 'Overridden', author: 'jdoc' } });
        const info = await pdfInfo(out);
        expect(info.metadata?.title).toBe('Overridden');
        expect(info.metadata?.author).toBe('jdoc');
    });

    it('未传 fonts 时中文叠加失败（标准字体无 CJK 字形），不静默空白', async () => {
        const src = await makeSource();
        await expect(
            editPdf(src, { overlays: [[{ type: 'text', x: 20, y: 300, size: 14, text: '中文签名' }]] })
        ).rejects.toThrow(/pdf/);
    });

    it('删光所有页时报错', async () => {
        const src = await makeSource();
        await expect(editPdf(src, { removePages: [0, 1, 2] })).rejects.toThrow(/无页面/);
    });

    it('非 PDF 输入报格式错误', async () => {
        await expect(editPdf(new Uint8Array([1, 2, 3]), {})).rejects.toThrow(/pdf/);
    });
});

describe('pdf：合并与拆分', () => {
    it('合并两份 PDF 按顺序拼接', async () => {
        const out = await mergePdfs([await makeSource(), await makeSource()]);
        expect((await pdfInfo(out)).pageCount).toBe(6);
    });

    it('空输入报「没有可合并的 PDF」', async () => {
        await expect(mergePdfs([])).rejects.toThrow(/没有可合并/);
    });

    it('按区间拆分为多个文件', async () => {
        const parts = await splitPdf(await makeSource(), [
            [1, 1],
            [2, 3]
        ]);
        expect(parts).toHaveLength(2);
        expect((await pdfInfo(parts[0])).pageCount).toBe(1);
        expect((await pdfInfo(parts[1])).pageCount).toBe(2);
        expect(await pdfToText(parts[0])).toContain('Page One');
        expect(await pdfToText(parts[1])).toContain('Page Two');
    });

    it('拆分区间越界时收敛到有效范围', async () => {
        const parts = await splitPdf(await makeSource(), [[2, 99]]);
        expect(parts).toHaveLength(1);
        expect((await pdfInfo(parts[0])).pageCount).toBe(2);
    });
});