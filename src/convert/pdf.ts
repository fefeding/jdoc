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
import { createAsset, createAssetId, type DocAsset } from '../core/asset';
import type { PdfImage, PdfMetadata, PdfTextItem, PdfVector } from '../pdf/types';
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
    /**
     * 是否还原页面图像，默认 true
     *
     * PDF 没有「块」概念，图文都是绝对定位的，图像必须靠坐标插回阅读顺序。
     * 关掉可省掉图像解码与 PNG 编码的开销（纯文本批量抽取场景）。
     */
    images?: boolean;
    /**
     * 单张图的像素上限，超出后降采样（默认 2e6，约 1600x1250）；0 表示不降采样
     *
     * 扫描件整页图常达 3000x4000，原样编码会让 md / html 膨胀到几百 MB。
     */
    maxImagePixels?: number;
    /**
     * pdfjs worker 地址（Node 端）。pdfjs 被打包器内联进 bundle 时，
     * 它内置的相对 worker 路径会失效，必须显式给出可加载的地址。
     */
    workerSrc?: string;
}

/** PDF -> blocks 结果 */
export interface PdfToBlocksResult {
    blocks: DocxBlock[];
    /** 文档元数据（若有） */
    metadata?: PdfMetadata;
    /**
     * 页面图像资源（图片块只存 assetId，字节在这里按 id 取）
     *
     * 转 md / html 时要把这张表一起传下去，否则图片解析不出 data URL 会整块丢弃。
     */
    assets?: DocAsset[];
}

/** 带纵向位置的图片块（用于按阅读顺序插回文本流） */
interface PlacedImage {
    /** 图像下边缘 y（PDF 坐标系，y 越大越靠上） */
    y: number;
    block: DocxBlock;
}

/** 一行文本（已含空格估算） */
interface Line {
    y: number;
    x: number;
    size: number;
    items: PdfTextItem[];
    text: string;
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
    const wantImages = options.images !== false;
    const { document } = await parsePdf(input, {
        items: true,
        vectors: quality === 'layout',
        text: false,
        images: wantImages,
        maxImagePixels: options.maxImagePixels,
        pages: options.pages,
        maxPages: options.maxPages,
        workerSrc: options.workerSrc
    });

    const allBlocks: DocxBlock[] = [];
    const assets: DocAsset[] = [];

    for (const page of document.pages || []) {
        const items = page.items || [];
        const pageImages = wantImages ? page.images || [] : [];

        // 纯图片页（扫描件、整页插图）：没有文本也要把图产出来
        if (!items.length) {
            if (!pageImages.length) continue;
            allBlocks.push(
                ...[...pageImages]
                    .sort((a, b) => b.y - a.y)
                    .map((img) => imageBlockOf(img, page.index, assets))
                    .filter((b): b is DocxBlock => !!b)
            );
            if (options.pageBreaks && page.index < (document.pageCount || 1)) {
                allBlocks.push({ type: 'pageBreak' });
            }
            continue;
        }

        // 正文字号：取出现频率最高者（直方图众数）
        const bodySize = modeSize(items.map((i) => i.fontSize)).size;
        const lineGroups = splitColumns(items, page.width);
        // 图像按 x 归属到栏，保证双栏排版时图跟着所在栏的文字走
        const imageGroups = assignImagesToGroups(pageImages, lineGroups);

        const pageBlocks: DocxBlock[] = [];
        for (let g = 0; g < lineGroups.length; g++) {
            const lines = groupLines(lineGroups[g], bodySize);
            const placed: PlacedImage[] = [];
            for (const img of imageGroups[g]) {
                const block = imageBlockOf(img, page.index, assets);
                if (block) placed.push({ y: img.y, block });
            }
            if (quality === 'layout' && page.vectors?.length) {
                extractTables(lines, page.vectors, pageBlocks);
            }
            pageBlocks.push(...linesToBlocks(lines, bodySize, placed));
        }
        allBlocks.push(...pageBlocks);
        // 非末页才补分页符，避免文档结尾多一个空页标记
        // 注意 parse.ts 的 page.index 是 1-based（1..pageCount）
        if (options.pageBreaks && page.index < (document.pageCount || 1)) {
            allBlocks.push({ type: 'pageBreak' });
        }
    }

    return { blocks: allBlocks, metadata: document.metadata, assets: assets.length ? assets : undefined };
}

/** 页面图像 -> 图片块（字节进资源表，块上只留 assetId 引用） */
function imageBlockOf(image: PdfImage, pageIndex: number, assets: DocAsset[]): DocxBlock | undefined {
    if (!image.data?.length) return undefined;
    const assetId = createAssetId('pdf', assets.length + 1);
    const asset = createAsset({
        id: assetId,
        data: image.data,
        mime: image.mime || 'image/png',
        part: `Page${pageIndex}:${image.id}`,
        anchor: {
            scope: 'pdf',
            kind: 'page',
            index: pageIndex,
            path: `pages[${pageIndex - 1}]`,
            label: `第 ${pageIndex} 页`
        },
        width: image.pixelWidth,
        height: image.pixelHeight
    });
    assets.push(asset);
    return {
        type: 'image',
        src: { assetId },
        assetId,
        // DocxImage 的宽高单位是磅，PDF 坐标本来就是磅，直接沿用
        width: image.width,
        height: image.height,
        alt: `第 ${pageIndex} 页图片`
    };
}

/** 图像按横向位置归属到栏（单栏时全部归该栏） */
function assignImagesToGroups(images: PdfImage[], groups: PdfTextItem[][]): PdfImage[][] {
    const out: PdfImage[][] = groups.map(() => []);
    if (!images.length) return out;
    if (groups.length === 1) {
        out[0] = images.slice();
        return out;
    }
    const ranges = groups.map((group) => {
        let min = Infinity;
        let max = -Infinity;
        for (const item of group) {
            min = Math.min(min, item.x);
            max = Math.max(max, item.x + item.width);
        }
        return { min, max };
    });
    for (const image of images) {
        const cx = image.x + image.width / 2;
        let best = 0;
        let bestDist = Infinity;
        ranges.forEach((range, index) => {
            const dist = cx < range.min ? range.min - cx : cx > range.max ? cx - range.max : 0;
            if (dist < bestDist) {
                bestDist = dist;
                best = index;
            }
        });
        out[best].push(image);
    }
    return out;
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
            line = { y: item.y, x: item.x, size: item.fontSize, items: [], text: '' };
            lines.push(line);
        }
        line.items.push(item);
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

/** 行的右边界（文本实际排到的位置，用来判断该行是否排到了行尾） */
function lineRight(line: Line): number {
    let right = 0;
    for (const it of line.items) right = Math.max(right, it.x + it.width);
    return right;
}

/** 折行拼接：中文之间不加空格，英文/数字之间保留空格 */
function joinWrapped(prev: string, next: string): string {
    const cjk = /[㐀-鿿豈-﫿＀-￯]/;
    if (!prev || !next) return prev + next;
    if (cjk.test(prev[prev.length - 1]) && cjk.test(next[0])) return prev + next;
    return `${prev} ${next}`;
}

/** 把行序列转成块（标题/列表/段落，并按 y 插入图片块） */
function linesToBlocks(lines: Line[], bodySize: number, images: PlacedImage[] = []): DocxBlock[] {
    const blocks: DocxBlock[] = [];
    let paraBuffer: string[] = [];
    let listBuffer: string[] = [];
    // 列表项折行后第二行没有编号/符号，会被当成普通段落拆出去，结果每个 <li>
    // 各自变成一个 <ol>，编号全部从 1 重新开始。用「上一行排到行尾 + 续行与列表
    // 项同列（未缩进）」识别折行，并回并进上一个列表项。
    const rightEdge = lines.reduce((max, line) => Math.max(max, lineRight(line)), 0);
    const fullLine = (line: Line) => lineRight(line) >= rightEdge - bodySize * 2;
    let listStartX = 0;
    let listLastFull = false;

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

    // 图像按「下边缘 y」从大到小（自上而下）排队：行也是自上而下遍历，
    // 遇到某行时把它上方的图先输出，图文顺序就与版面一致。
    const pending = [...images].sort((a, b) => b.y - a.y);
    let nextImage = 0;
    const flushImagesAbove = (y: number) => {
        while (nextImage < pending.length && pending[nextImage].y > y) {
            // 图文之间必须断开：否则图片会被并进正在攒的段落
            flushPara();
            flushList();
            blocks.push(pending[nextImage].block);
            nextImage++;
        }
    };

    for (const line of lines) {
        flushImagesAbove(line.y);
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
            listStartX = line.x;
            listLastFull = fullLine(line);
        } else if (
            listBuffer.length &&
            listLastFull &&
            // 续行与列表项同列：缩进行（目录子条目等）是新内容，不并进列表项
            line.x <= listStartX + Math.max(1.5, bodySize * 0.3)
        ) {
            listBuffer[listBuffer.length - 1] = joinWrapped(listBuffer[listBuffer.length - 1], line.text);
            listLastFull = fullLine(line);
        } else {
            flushList();
            paraBuffer.push(line.text);
        }
    }
    // 行走完后剩下的是排在末尾的图（页脚上方 / 文末插图）
    flushImagesAbove(-Infinity);
    flushPara();
    flushList();
    return blocks;
}

/** 用矢量线还原表格，把表格区域内的文本行转成表格块 */
function extractTables(lines: Line[], vectors: PdfVector[], out: DocxBlock[]): void {
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
    // 注意坐标系：PDF 的 y 向上为正，hY 升序数组里 hY[0] 是最靠下的横线，
    // 因此 rows[0] 对应页面最底部的行带，构建完后反转才是从上到下的阅读顺序
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
        rows.reverse();
        // 纯空白字符的单元格归一为空串（pdfjs 可能拆出独立空格 item）
        const cleaned = rows.map((row) => row.map((cell) => cell.trim()));
        out.push({ type: 'table', header: true, rows: cleaned });
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
