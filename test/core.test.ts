import { describe, expect, it } from 'vitest';
import { toArrayBuffer, toUint8Array, bytesToBase64 } from '../src/core/input';
import { parseXml, findFirst, findAll, attr, textOf, childElements, firstChild } from '../src/core/xml';
import { detectFormat } from '../src/index';

describe('core/input：入参归一化', () => {
    const payload = new Uint8Array([1, 2, 3, 4, 250]);

    it('接受 Uint8Array', async () => {
        expect(await toUint8Array(payload)).toEqual(payload);
    });

    it('接受 ArrayBuffer', async () => {
        const buffer = payload.buffer.slice(0) as ArrayBuffer;
        expect(await toUint8Array(buffer)).toEqual(payload);
    });

    it('接受 TypedArray 视图并保留偏移', async () => {
        const backing = new Uint8Array([9, 9, 1, 2, 3, 4, 250, 9]);
        const view = backing.subarray(2, 7);
        const result = await toUint8Array(view);
        expect(Array.from(result)).toEqual([1, 2, 3, 4, 250]);
    });

    it('接受 base64 字符串', async () => {
        const base64 = bytesToBase64(payload);
        expect(Array.from(await toUint8Array(base64))).toEqual(Array.from(payload));
    });

    it('接受 data URL', async () => {
        const url = `data:application/octet-stream;base64,${bytesToBase64(payload)}`;
        expect(Array.from(await toUint8Array(url))).toEqual(Array.from(payload));
    });

    it('Node 下精确切分 Buffer（不带上相邻内存）', async () => {
        const backing = Buffer.from([7, 7, 1, 2, 3, 4, 250, 7]);
        const slice = backing.subarray(2, 7);
        const result = await toUint8Array(slice);
        expect(result.byteLength).toBe(5);
        expect(Array.from(result)).toEqual([1, 2, 3, 4, 250]);
    });

    it('toArrayBuffer 长度与内容一致', async () => {
        const buffer = await toArrayBuffer(payload);
        expect(buffer.byteLength).toBe(payload.length);
        expect(Array.from(new Uint8Array(buffer))).toEqual(Array.from(payload));
    });

    it('空输入抛错', async () => {
        await expect(toUint8Array(null as never)).rejects.toThrow(TypeError);
    });
});

describe('core/xml：XML 解析', () => {
    const xml = `<?xml version="1.0"?>
<root xmlns:w="urn:w">
  <!-- 注释 -->
  <w:item id="1" name="a&amp;b">文本&lt;1&gt;</w:item>
  <w:item id="2"><![CDATA[raw <not> parsed]]></w:item>
  <w:empty/>
</root>`;

    it('跳过声明与注释并保留元素', () => {
        const roots = parseXml(xml);
        expect(roots.length).toBe(1);
        expect(roots[0].type).toBe('element');
    });

    it('解码命名实体与数字实体', () => {
        const item = findAll(parseXml(xml), 'item')[0];
        expect(textOf(item)).toBe('文本<1>');
        expect(attr(item, 'name')).toBe('a&b');
    });

    it('CDATA 原样保留且不解码实体', () => {
        const items = findAll(parseXml(xml), 'item');
        expect(textOf(items[1])).toBe('raw <not> parsed');
    });

    it('属性名兼容无前缀写法', () => {
        const item = findAll(parseXml(xml), 'item')[0];
        expect(attr(item, 'id')).toBe('1');
        expect(attr(item, 'styleId')).toBeUndefined();
    });

    it('findFirst / childElements / firstChild 可用', () => {
        const root = parseXml(xml)[0];
        expect(findFirst(root, 'item')).toBeDefined();
        expect(childElements(root, 'item').length).toBe(2);
        expect(firstChild(root, 'empty')).toBeDefined();
    });

    it('容错未闭合标签', () => {
        const broken = parseXml('<a><b>x</a>');
        expect(broken.length).toBe(1);
    });
});

describe('detectFormat：格式识别', () => {
    it('识别 PDF', async () => {
        expect(await detectFormat(new TextEncoder().encode('%PDF-1.7\n'))).toBe('pdf');
    });

    it('识别 ZIP 类容器（依据内部标记部件）', async () => {
        const fake = new TextEncoder().encode('PK\u0003\u0004word/document.xml');
        expect(await detectFormat(fake)).toBe('docx');
        expect(await detectFormat(new TextEncoder().encode('PK\u0003\u0004xl/workbook.xml'))).toBe('xlsx');
        expect(await detectFormat(new TextEncoder().encode('PK\u0003\u0004ppt/presentation.xml'))).toBe('pptx');
    });

    it('无法识别时返回 null', async () => {
        expect(await detectFormat(new TextEncoder().encode('random bytes here'))).toBeNull();
    });
});
