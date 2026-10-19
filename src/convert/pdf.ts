/**
 * PDF -> 文档块模型（DocxBlock）解析（带版面分析）
 *
 * 基于 pdfjs-dist 提取的文本片段（含坐标/字号）与矢量线段，做启发式版面还原：
 *
 * - 行分组：按基线 y 聚类、行内按 x 排序
 * - 标题识别：字号相对正文显著偏大（或加粗）且较短
 * - 列表识别：行首项目符号 / 有序编号
 * - 双栏：检测页面中部的纵向空白走廊，把内容切成左右两栏（先左后右）
 * - 表格（layout 模式）：用矢量线聚合成网格，把落在单元格内的文本还原成表格
 *
 * 这是「有损但可用」的还原：复杂排版（绕排、跨页表、竖排）无法完美还原，
 * 但标题/段落/列表/双栏论文/带边框表格等常见结构都能正确拆分。
 *
 * @module convert/pdf
 */

import { parsePdf } from '../pdf/parse';
import type { PdfMetadata, PdfTextItem, PdfVector } from '../pdf/types';
import type { DocxBlock } from '../docx/types';

/** PDF -> blocks 选项 */
export interface PdfToBlocksOptions {
    /** 还原质量：'fast' 仅行分组 + 标题/列表；'layout' 另加双栏与表格还原。默认 'layout' */
    quality?: 'fast' | 'layout';
    /** 只解析前 N 页 */
    maxPages?: number;
    /** 只解析指定页（1 起） */
    pages?: number[];
    /** 是否在各页之间插入 pageBreak 块，默认 false */
    pageBreaks?: boolean;
}

/** PDF -> blocks 结果 */
export interface PdfToBlocksResult {
    blocks: DocxBlock[];
    /** 文档元数据（若有） */
    metadata?: PdfMetadata;
}

/** 一行文本（已含空格估算） */
interface Line {
    y: number;
    x: number;
    size: number;
    items: PdfTextItem[];
    text: string;
    bold: boolean;
}

/** 行首列表符号/编号 */
const LIST_LEAD = /^\s*(?:[-•·◦▪●*‣◦]|\(?[0-9]{1,3}[.、)）]|[a-zA-Z][.、)）])\s+/;

/**
 * 把 PDF 解析为文档块模型
 * @param input - PDF 字节 / data URL / base64 / Blob
 * @param options - 版面分析选项
 */
export async function pdfToBlocks(
    input: Parameters<typeof parsePdf>[0],
    options: PdfToBlocksOptions = {}
): Promise<PdfToBlocksResult> {
    const quality = options.quality ?? 'layout';
    const { document } = await parsePdf(input, {
        items: true,
        vectors: quality === 'layout',
        text: false,
        pages: options.pages,
        maxPages: options.maxPages
    });

    const allBlocks: DocxBlock[] = [];

    for (const page of document.pages || []) {
        const items = page.items || [];
        if (!items.length) continue;

        // 正文字号：取出现频率最高者（直方图众数）
        const bodySize = modeSize(items.map((i) => i.fontSize)).size;
        const lineGroups = splitColumns(items, page.width);

        const pageBlocks: DocxBlock[] = [];
        for (const group of lineGroups) {
            const lines = groupLines(group, bodySize);
            if (quality === 'layout' && page.vectors?.length) {
                extractTables(lines, page.vectors, bodySize, pageBlocks);
            }
            pageBlocks.push(...linesToBlocks(lines, bodySize));
        }
        allBlocks.push(...pageBlocks);
        if (options.pageBreaks && page.index < (document.pageCount || 0)) {
            allBlocks.push({ type: 'pageBreak' });
        }
    }

    return { blocks: allBlocks, metadata: document.metadata };
}

// ---------------------------------------------------------------------------
// 行 / 栏 / 块 分析
// ---------------------------------------------------------------------------

/** 按正文大小分组连续的行 */
function groupLines(items: PdfTextItem[], bodySize: number): Line[] {
    const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
    const lines: Line[] = [];
    const tol = Math.max(1.5, bodySize * 0.6);
    for (const item of sorted) {
        let line = lines.find((l) => Math.abs(l.y - item.y) <= tol && Math.abs(l.y - item.y) <= Math.max(tol, l.size * 0.6));
        if (!line) {
            line = { y: item.y, x: item.x, size: item.fontSize, items: [], text: '', bold: false };
            lines.push(line);
        }
        line.items.push(item);
        if (/bold/i.test(item.fontName || '')) line.bold = true;
        line.size = Math.max(line.size, item.fontSize);
    }
    // 行内按 x 排序并拼出文本（依据 x 间隔判断是否需要空格）
    for (const line of lines) {
        line.items.sort((a, b) => a.x - b.x);
        let text = '';
        let prevEnd = -Infinity;
        for (const it of line.items) {
            const isCjk = /[一-鿿豈-﫿＀-￯]/.test(it.text);
            const gap = it.x - prevEnd;
            if (prevEnd > -Infinity && gap > it.fontSize * 0.25 && !isCjk && !/^\s/.test(it.text)) text += ' ';
            text += it.text;
            prevEnd = it.x + it.width;
        }
        line.text = text.trim();
        line.x = line.items[0].x;
    }
    lines.sort((a, b) => b.y - a.y);
    return lines;
}

/**
 * 双栏检测：在页面中部找一条纵向空白走廊，把内容切成左右两栏。
 * 返回按「先左后右」顺序的两个 item 分组（单栏时返回整页一组）。
 */
function splitColumns(items: PdfTextItem[], pageWidth: number): PdfTextItem[][] {
    const xs = items.map((i) => i.x + i.width / 2);
    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const span = maxX - minX;
    if (span < pageWidth * 0.5) return [items];

    // 在 [20%, 80%] 区间内找最宽的空白带
    const step = span / 100;
    let bestGap = 0;
    let bestAt = -1;
    for (let gx = minX + span * 0.2; gx < maxX - span * 0.2; gx += step) {
        const occupied = xs.some((x) => Math.abs(x - gx) < span * 0.04);
        if (occupied) continue;
        // 向两侧扩展空白
        let left = gx;
        let right = gx;
        while (left > minX && !xs.some((x) => Math.abs(x - left) < span * 0.04)) left -= step;
        while (right < maxX && !xs.some((x) => Math.abs(x - right) < span * 0.04)) right += step;
        const gap = right - left;
        if (gap > bestGap) {
            bestGap = gap;
            bestAt = (left + right) / 2;
        }
    }
    if (bestGap > span * 0.12 && bestAt > 0) {
        const left = items.filter((i) => i.x + i.width / 2 < bestAt);
        const right = items.filter((i) => i.x + i.width / 2 >= bestAt);
        if (left.length && right.length) return [left, right];
    }
    return [items];
}

/** 把行序列转成块（标题/列表/段落） */
function linesToBlocks(lines: Line[], bodySize: number): DocxBlock[] {
    const blocks: DocxBlock[] = [];
    let paraBuffer: string[] = [];
    let listBuffer: string[] = [];

    const flushPara = () => {
        const text = paraBuffer.join(' ').replace(/\s+/g, ' ').trim();
        if (text) blocks.push({ type: 'paragraph', text });
        paraBuffer = [];
    };
    const flushList = () => {
        if (!listBuffer.length) return;
        blocks.push({
            type: 'list',
            ordered: /^\s*\d/.test(listBuffer[0]),
            items: listBuffer.map((t) => {
                const m = LIST_LEAD.exec(t);
                return (m ? t.slice(m[0].length) : t) as never;
            })
        });
        listBuffer = [];
    };

    for (const line of lines) {
        const isHeading = line.size > bodySize * 1.18 && line.text.length <= Math.max(60, bodySize * 6);
        const isList = LIST_LEAD.test(line.text);

        if (isHeading) {
            flushPara();
            flushList();
            const ratio = line.size / bodySize;
            const level = ratio >= 1.6 ? 1 : ratio >= 1.35 ? 2 : ratio >= 1.2 ? 3 : 4;
            blocks.push({ type: 'heading', level: level as 1 | 2 | 3 | 4 | 5 | 6, text: line.text });
        } else if (isList) {
            flushPara();
            listBuffer.push(line.text);
        } else {
            flushList();
            paraBuffer.push(line.text);
        }
    }
    flushPara();
    flushList();
    return blocks;
}

/** 用矢量线还原表格，把表格区域内的文本行转成表格块 */
function extractTables(lines: Line[], vectors: PdfVector[], _bodySize: number, out: DocxBlock[]): void {
    const hLines = vectors.filter((v) => v.horizontal && Math.abs(v.x2 - v.x1) > 10);
    const vLines = vectors.filter((v) => v.vertical && Math.abs(v.y2 - v.y1) > 10);
    if (hLines.length < 2 || vLines.length < 2) return;

    // 按 y / x 聚类
    const hY = cluster(hLines.map((v) => (v.y1 + v.y2) / 2));
    const vX = cluster(vLines.map((v) => (v.x1 + v.x2) / 2));
    if (hY.length < 2 || vX.length < 2) return;

    const minX = Math.min(...vX) - 2;
    const maxX = Math.max(...vX) + 2;
    const minY = Math.min(...hY) - 2;
    const maxY = Math.max(...hY) + 2;

    // 收集落在表格框内的行
    const inside = lines.filter((l) => {
        const cy = l.y;
        const cx = l.x + (l.items[0]?.width || 0) / 2;
        return cy <= maxY && cy >= minY && cx >= minX && cx <= maxX;
    });
    if (!inside.length) return;

    // 每个行 -> 行带；每个 item -> 列带
    const rows: string[][] = hY.slice(0, -1).map(() => vX.slice(0, -1).map(() => ''));
    for (const line of inside) {
        const r = hY.findIndex((y) => y > line.y) - 1;
        if (r < 0) continue;
        for (const item of line.items) {
            const cx = item.x + item.width / 2;
            const c = vX.findIndex((x) => x > cx) - 1;
            if (c < 0 || c >= rows[r].length) continue;
            rows[r][c] += (rows[r][c] ? ' ' : '') + item.text;
        }
    }
    if (rows.length && rows[0].length) {
        out.push({ type: 'table', header: true, rows: rows as never });
    }
    // 从行集合中移除已归入表格的文本，避免重复
    for (const line of inside) {
        const idx = lines.indexOf(line);
        if (idx >= 0) lines.splice(idx, 1);
    }
}

/** 按容差聚类一维坐标 */
function cluster(values: number[], tol = 4): number[] {
    const sorted = [...values].sort((a, b) => a - b);
    const groups: number[][] = [];
    for (const v of sorted) {
        const last = groups[groups.length - 1];
        if (last && v - last[last.length - 1] <= tol) last.push(v);
        else groups.push([v]);
    }
    return groups.map((g) => g.reduce((s, x) => s + x, 0) / g.length);
}

/** 字号众数（出现最频繁的大小） */
function modeSize(sizes: number[]): { size: number; count: number } {
    const hist = new Map<number, number>();
    for (const s of sizes) {
        const key = Math.round(s * 2) / 2;
        hist.set(key, (hist.get(key) || 0) + 1);
    }
    let best = 12;
    let bestCount = 0;
    for (const [size, count] of hist) {
        if (count > bestCount) {
            bestCount = count;
            best = size;
        }
    }
    return { size: best, count: bestCount };
}
