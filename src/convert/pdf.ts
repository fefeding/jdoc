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
import type { PdfImage, PdfMetadata, PdfPage, PdfTextItem, PdfVector } from '../pdf/types';
import type { DocxBlock, DocxRun } from '../docx/types';

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
    /** cMap / 标准字体 / wasm 资源目录与内嵌资源 worker（Node 端，透传给 parsePdf） */
    cMapUrl?: string;
    standardFontDataUrl?: string;
    wasmUrl?: string;
    embeddedWorkerPath?: string;
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
    /** 该行主字体名（pdfjs 对内嵌子集字体会改成 g_d0_f2 之类，命中率低，仅作辅助信号） */
    font?: string;
    /** 该行是表格的表头行（rows 为整张表），由表格探测标记 */
    table?: { rows: string[][] };
}

/** 行内代码使用的等宽字体名（docx/HTML 都按它出等宽效果） */
const CODE_FONT = 'Consolas';

/** 合并相邻的同样式 run */
function mergeRuns(runs: DocxRun[]): DocxRun[] {
    const out: DocxRun[] = [];
    for (const r of runs) {
        if (!r.text) continue;
        const last = out[out.length - 1];
        const same =
            !!last &&
            !!last.bold === !!r.bold &&
            !!last.italic === !!r.italic &&
            !!last.code === !!r.code &&
            (last.font ?? '') === (r.font ?? '');
        if (same) last.text += r.text;
        else out.push({ ...r });
    }
    return out;
}

/** 从 run 序列头部去掉 n 个字符（用于剥掉列表符号） */
function stripRunPrefix(runs: DocxRun[], n: number): DocxRun[] {
    if (n <= 0) return runs;
    const out = runs.map((r) => ({ ...r }));
    let remain = n;
    for (const r of out) {
        if (remain <= 0) break;
        if (r.text.length <= remain) {
            remain -= r.text.length;
            r.text = '';
        } else {
            r.text = r.text.slice(remain);
            remain = 0;
        }
    }
    return out.filter((r) => r.text);
}

/**
 * 正文字号：文档级统计，取「含中文片段」的字号众数。
 *
 * 代码多为 ASCII，若直接取全部片段的众数，代码密集的文档会把代码字号
 * （实测 10pt）当成正文（11pt），导致标题判定失准、代码行认不出来。
 */
function documentBodySize(pages: PdfPage[]): number {
    const cjk = new Map<number, number>();
    const all = new Map<number, number>();
    // 各字号覆盖的字符数：出现次数相同时用它决胜
    const chars = new Map<number, number>();
    for (const page of pages) {
        for (const item of page.items || []) {
            if (!item.text?.trim()) continue;
            const key = Math.round(item.fontSize * 2) / 2;
            all.set(key, (all.get(key) || 0) + 1);
            chars.set(key, (chars.get(key) || 0) + item.text.trim().length);
            if (/[一-鿿㐀-䶿豈-﫿]/.test(item.text)) cjk.set(key, (cjk.get(key) || 0) + 1);
        }
    }
    const pick = (hist: Map<number, number>): number => {
        let best = 0;
        let bestCount = 0;
        let bestChars = 0;
        for (const [size, count] of hist) {
            const ch = chars.get(size) || 0;
            // 出现次数优先；次数相同时取覆盖字符更多的字号
            //（否则「一条大标题 + 一段正文」这种文档会把标题字号当成正文：
            //  标题行与正文行被并成同一行，标题层级彻底丢失）
            if (count > bestCount || (count === bestCount && ch > bestChars)) {
                best = size;
                bestCount = count;
                bestChars = ch;
            }
        }
        return best;
    };
    return pick(cjk.size ? cjk : all) || 12;
}

/**
 * 一行是否整体用更小字号（PDF 里代码块/文件树/注释块常用更小字号排版）。
 *
 * 注意不能用 line.size（它是该行最大字号）：正文里混排行内代码的行，
 * 最大字号仍是正文，只有「整行都是小字号」才是真正的代码块/小字块。
 */
function lineIsSmallSized(line: Line, bodySize: number): boolean {
    const items = line.items.filter((it) => it.text.trim());
    if (!items.length) return false;
    return items.every((it) => it.fontSize < bodySize - 0.4);
}

/**
 * 行的 run 序列：字号小于正文的片段判为行内代码。
 *
 * PDF 排版里行内代码（any、vue-tsc、prd-to-ui 等）通常与正文同字体但字号更小，
 * 识别出来后 md 会输出反引号、HTML/docx 会走等宽样式。
 */
function lineRuns(line: Line, bodySize: number): DocxRun[] {
    const runs: DocxRun[] = [];
    let prevEnd = -Infinity;
    for (const it of line.items) {
        if (!it.text) continue;
        const isCjk = /[一-鿿]/.test(it.text);
        const gap = it.x - prevEnd;
        let text = it.text;
        if (prevEnd > -Infinity && gap > it.fontSize * 0.25 && !isCjk && !/^\s/.test(text)) text = ` ${text}`;
        const small = it.fontSize < bodySize - 0.4;
        runs.push(small ? { text, code: true, font: CODE_FONT } : { text });
        prevEnd = it.x + it.width;
    }
    return mergeRuns(runs);
}

/**
 * 一行按列切分为单元格。
 *
 * 关键：PDF 把「列间隙」编码成一个独立的空白片段（文本是空格、宽度正好等于
 * 间隙宽度，如 w=159.5），所以不能靠 item 之间的 x 间隔判断换列——间隔已被
 * 这个空白片段吃掉。这里以「空白片段宽度 >= 阈值」作为换列信号。
 */
function splitCells(line: Line, gap: number): { x: number; text: string }[] {
    const cells: { x: number; text: string }[] = [];
    let cur: { x: number; end: number; text: string } | null = null;
    let pendingBreak = false;
    for (const it of line.items) {
        if (!it.text) continue;
        if (/^\s+$/.test(it.text)) {
            // 空白片段：宽度足够大 = 列间距，否则只是词间空格
            if (it.width >= gap) pendingBreak = true;
            continue;
        }
        if (pendingBreak || !cur) {
            cur = { x: it.x, end: it.x + it.width, text: it.text };
            cells.push(cur);
            pendingBreak = false;
            continue;
        }
        const gapToPrev = it.x - cur.end;
        const isCjk = /[一-鿿]/.test(it.text);
        let text = it.text;
        if (gapToPrev > it.fontSize * 0.25 && !isCjk && !/^\s/.test(text)) text = ` ${text}`;
        cur.text += text;
        cur.end = Math.max(cur.end, it.x + it.width);
    }
    return cells;
}

/**
 * 无边框表格还原：靠「相邻多行的列起点对齐」识别。
 *
 * 有框线表格走矢量线聚合（extractTables），但很多文档的表格没有框线、只有列对齐。
 * 这里要求：连续若干行列数相同、各列起点 x 对齐（容差 10pt）、且不是项目符号行。
 * 代码行虽然也会被切成多段，但列位置每行都在变，不会对齐，因此不会误判。
 */
function extractBorderlessTables(lines: Line[], bodySize: number): void {
    const gap = Math.max(24, bodySize * 2.5);
    const consumed = new Set<number>();
    let i = 0;
    while (i < lines.length) {
        if (consumed.has(i)) { i++; continue; }
        const line = lines[i];
        if (LIST_LEAD.test(line.text)) { i++; continue; }
        const cells = splitCells(line, gap);
        if (cells.length < 2) { i++; continue; }
        // 向后扩展：列数相同且各列起点对齐
        const run: number[] = [i];
        for (let j = i + 1; j < lines.length; j++) {
            if (consumed.has(j) || LIST_LEAD.test(lines[j].text)) break;
            const cj = splitCells(lines[j], gap);
            if (cj.length !== cells.length) break;
            let aligned = true;
            for (let k = 0; k < cj.length; k++) {
                if (Math.abs(cj[k].x - cells[k].x) > 10) { aligned = false; break; }
            }
            if (!aligned) break;
            run.push(j);
        }
        // 至少 3 行（表头 + 2 行）才认定为表格，避免两行巧合对齐被误判
        if (run.length < 3) { i++; continue; }
        // 表头首格像文件名（home.ts / app.json…）说明这是「文件树/结构清单」，
        // 只是两列排版整齐而已，按表格还原会切得支离破碎
        const head = (cells[0]?.text ?? '').trim();
        if (
            /(?:^|\s)[^\s]+\.(js|ts|tsx|jsx|json|vue|wxml|wxss|html|css|scss|less|md|txt|yml|yaml)\s*$/i.test(head) ||
            /^#/.test(head)
        ) {
            i = run[run.length - 1] + 1;
            continue;
        }
        const rows = run.map((idx) => splitCells(lines[idx], gap).map((c) => c.text.trim()));
        // 标记在首行上：交给 linesToBlocks 按 y 位置插入，否则表格会被塞到页首
        lines[run[0]].table = { rows };
        run.slice(1).forEach((idx) => {
            const pos = lines.indexOf(lines[idx]);
            if (pos >= 0) lines.splice(pos, 1);
        });
        i = run[run.length - 1] + 1;
    }
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
        workerSrc: options.workerSrc,
        cMapUrl: options.cMapUrl,
        standardFontDataUrl: options.standardFontDataUrl,
        wasmUrl: options.wasmUrl,
        embeddedWorkerPath: options.embeddedWorkerPath
    });

    const allBlocks: DocxBlock[] = [];
    const assets: DocAsset[] = [];

    // 正文字号：文档级统计（含中文片段的字号众数）。按页取众数时，代码密集页
    // 会把代码字号当成正文，导致标题层级判定失准、代码行认不出来。
    const docBodySize = documentBodySize(document.pages || []);

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

        // 正文字号：优先用文档级统计，兜底退回本页众数
        const bodySize = docBodySize || modeSize(items.map((i) => i.fontSize)).size;
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
                extractTables(lines, page.vectors);
            }
            // 无边框表格（只有列对齐、没有框线）：靠列起点对齐识别
            if (quality === 'layout') {
                extractBorderlessTables(lines, bodySize);
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
        line.font = line.font ?? item.fontName;
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

/** 一行是否「看起来像代码」 */
function lineLooksCode(line: Line, bodySize: number): boolean {
    const t = line.text;
    if (!t.trim()) return false;
    // 字号明显小于正文：PDF 排版代码常用更小字号（实测：正文 11pt / 代码 10pt）
    const small = lineIsSmallSized(line, bodySize);
    // 1) 等宽字体（内嵌子集字体名常被 pdfjs 改成 g_d0_f2，命中率低，仅作加分）
    if (line.font && /(mono|courier|consol|menlo|monospace|sfmono|liberation\s*mono)/i.test(line.font)) return true;
    // 2) 行内明显代码符号组合（强信号，避免 "{ ret, data, msg }" 这类中文夹符号被误判）
    if (/(=>|function\s|const\s|let\s|var\s|import\s|export\s|\basync\b|\bawait\b|\bclass\b|\breturn\b|\bpublic\b|\bprivate\b|module\.exports|\brequire\(|\bconsole\.|\bdef\s|\bSELECT\b|\bFROM\b|\bINSERT\b|\bUPDATE\b|\bDELETE\b)/i.test(t)) return true;
    // 3) 以常见代码关键字开头
    if (/^\s*(import|export|const|let|var|function|class|def|public|private|void|async|await|#include|#define|#import|SELECT|FROM|INSERT|UPDATE|DELETE)\b/i.test(t)) return true;
    // 4) shell 命令与注释行（npm/yarn/git/pip 等；# 注释需字号更小，避免正文井号误判）
    if (/^\s*(npm|yarn|pnpm|npx|pip3?|git|docker|curl|wget|cd|mkdir|cp|mv|rm|echo|cat|sudo|apt|brew)\s/.test(t)) return true;
    if (small && /^\s*#\s*\S/.test(t)) return true;
    // 5) 结构括号：行尾是开括号 / 行首是收括号——多行块的典型形态。
    // 这类信号与字号无关（很多文档代码与正文字号相同），正文很少这么写。
    if (/[{([]\s*$/.test(t.trim()) || /^\s*[})\]]/.test(t.trim())) return true;
    return false;
}

/** 代码块续行是否仍像代码（弱信号）：用于括号深度向前扩展时，避免把散文/表格行吞进代码块 */
function looksLikeCodeContinuation(line: Line): boolean {
    const t = line.text.trim();
    if (!t) return true; // 空行：保留在代码块内
    if (/^\s*[})\]]/.test(t)) return true; // 收尾括号
    if (/[{([]\s*$/.test(t)) return true; // 开头括号（块首行续）
    if (/\/\/|\/\*/.test(t)) return true; // 注释
    if (/["'`]/.test(t)) return true; // 字符串
    if (/[,;]\s*$/.test(t)) return true; // 行尾逗号/分号（对象/数组项）
    if (/\b[a-zA-Z_$][\w$]*\s*[:=]/.test(t)) return true; // key: / var =
    if (/\//.test(t)) return true; // 路径（文件树：pages/、app.js / app.ts）
    if (/#/.test(t)) return true; // 井号注释（文件树说明：# 页面逻辑）
    return false;
}

/** 把行序列转成块（标题/列表/段落/代码块，并按 y 插入图片块） */
function linesToBlocks(lines: Line[], bodySize: number, images: PlacedImage[] = []): DocxBlock[] {
    const blocks: DocxBlock[] = [];
    // 段落与列表都按 run 级保存：字号小于正文的片段会标成行内代码
    let paraBuffer: DocxRun[][] = [];
    let listBuffer: DocxRun[][] = [];
    // 列表项折行后第二行没有编号/符号，会被当成普通段落拆出去，结果每个 <li>
    // 各自变成一个 <ol>，编号全部从 1 重新开始。用「上一行排到行尾 + 续行与列表
    // 项同列（未缩进）」识别折行，并回并进上一个列表项。
    const rightEdge = lines.reduce((max, line) => Math.max(max, lineRight(line)), 0);
    const fullLine = (line: Line) => lineRight(line) >= rightEdge - bodySize * 2;
    let listStartX = 0;
    let listLastFull = false;

    const flushPara = () => {
        const runs: DocxRun[] = [];
        paraBuffer.forEach((lineRunList, idx) => {
            if (idx > 0) runs.push({ text: ' ' });
            runs.push(...lineRunList);
        });
        const merged = mergeRuns(runs);
        const text = merged.map((r) => r.text).join('').replace(/\s+/g, ' ').trim();
        if (text) blocks.push({ type: 'paragraph', text, runs: merged });
        paraBuffer = [];
    };
    const flushList = () => {
        if (!listBuffer.length) return;
        const firstText = listBuffer[0].map((r) => r.text).join('');
        blocks.push({
            type: 'list',
            ordered: /^\s*\d/.test(firstText),
            items: listBuffer.map((runs) => {
                const joined = runs.map((r) => r.text).join('');
                const m = LIST_LEAD.exec(joined);
                return (m ? stripRunPrefix(runs, m[0].length) : mergeRuns(runs)) as never;
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

    // 代码块识别：先判每行是否像代码，再用括号深度向前扩展（如 `module.exports = {`
    // 之后的 `printWidth: 100,` 没有关键字，靠未闭合的 { 判断仍属同一代码块）。
    // 仅当首行是强信号时才展开，避免正文中孤立的 { 误带出后续行。
    const codeFlag: boolean[] = new Array(lines.length).fill(false);
    {
        let depth = 0;
        let inCodeBlock = false;
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            // 整行字号比正文小一档的通常是代码/文件树（PDF 排版惯例）
            const small = lineIsSmallSized(line, bodySize);
            if (lineLooksCode(line, bodySize)) {
                codeFlag[i] = true;
                inCodeBlock = true;
            } else if (small && looksLikeCodeContinuation(line)) {
                codeFlag[i] = true;
                inCodeBlock = true;
            } else if (inCodeBlock && depth > 0 && looksLikeCodeContinuation(line)) {
                codeFlag[i] = true;
            } else {
                // 任何非代码续行立即结束代码上下文，防止括号深度跨页泄漏吞掉后续散文/表格
                inCodeBlock = false;
            }
            if (codeFlag[i]) {
                const open = (line.text.match(/[{([]/g) || []).length;
                const close = (line.text.match(/[})\]]/g) || []).length;
                depth += open - close;
            }
            if (depth <= 0) {
                depth = 0;
                inCodeBlock = false;
            }
        }
    }

    let codeRunLines: Line[] = [];
    const flushCode = () => {
        if (codeRunLines.length >= 2) {
            // 按等宽字宽（≈ 0.5 × 字号）还原块内相对缩进
            const minX = Math.min(...codeRunLines.map((l) => l.x));
            const text = codeRunLines
                .map((l) => {
                    const indent = Math.max(0, Math.round((l.x - minX) / (l.size * 0.5)));
                    return `${' '.repeat(indent)}${l.text}`;
                })
                .join('\n');
            blocks.push({ type: 'code', text });
        } else if (codeRunLines.length === 1) {
            // 单行代码不成立代码块（会丢内容），退回普通段落
            const only = codeRunLines[0];
            const runs = lineRuns(only, bodySize);
            const text = runs.map((r) => r.text).join('').trim();
            if (text) blocks.push({ type: 'paragraph', text, runs });
        }
        codeRunLines = [];
    };

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        flushImagesAbove(line.y);
        // 表格锚点行：按原位输出（否则表格会被塞到页首，顺序全乱）
        if (line.table) {
            flushPara();
            flushList();
            blocks.push({ type: 'table', header: true, rows: line.table.rows });
            continue;
        }
        if (codeFlag[i]) {
            // 代码块内：空行若前后都是代码则保留，否则忽略（不拆块）
            if (!line.text.trim() && codeRunLines.length && i + 1 < lines.length && codeFlag[i + 1]) {
                codeRunLines.push({ ...line, text: '' });
            } else if (line.text.trim()) {
                codeRunLines.push(line);
            }
            continue;
        }
        flushCode();

        const isHeading = line.size > bodySize * 1.18 && line.text.length <= Math.max(60, bodySize * 6);
        const isList = LIST_LEAD.test(line.text);
        const runsOfLine = lineRuns(line, bodySize);

        if (isHeading) {
            flushPara();
            flushList();
            const ratio = line.size / bodySize;
            const level = ratio >= 1.6 ? 1 : ratio >= 1.35 ? 2 : ratio >= 1.2 ? 3 : 4;
            blocks.push({
                type: 'heading',
                level: level as 1 | 2 | 3 | 4 | 5 | 6,
                text: line.text,
                runs: runsOfLine
            });
        } else if (isList) {
            flushPara();
            listBuffer.push(runsOfLine);
            listStartX = line.x;
            listLastFull = fullLine(line);
        } else if (
            listBuffer.length &&
            listLastFull &&
            // 续行与列表项同列：缩进行（目录子条目等）是新内容，不并进列表项
            line.x <= listStartX + Math.max(1.5, bodySize * 0.3)
        ) {
            const last = listBuffer[listBuffer.length - 1];
            const prevText = last.map((r) => r.text).join('');
            // 中文之间不补空格，其余情况补一个（与旧的 joinWrapped 语义一致）
            const needSpace = !(
                /[㐀-鿿豈-﫿＀-￯]/.test(prevText.slice(-1)) && /[㐀-鿿豈-﫿＀-￯]/.test(line.text.slice(0, 1))
            );
            if (needSpace) last.push({ text: ' ' });
            last.push(...runsOfLine);
            listLastFull = fullLine(line);
        } else {
            flushList();
            paraBuffer.push(runsOfLine);
        }
    }
    // 行走完后剩下的是排在末尾的图（页脚上方 / 文末插图）
    flushImagesAbove(-Infinity);
    flushCode();
    flushPara();
    flushList();
    return blocks;
}

/** 用矢量线还原表格，把表格区域内的文本行转成表格块 */
function extractTables(lines: Line[], vectors: PdfVector[]): void {
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
        // 装饰性网格守卫：封面/页面的装饰线会被聚合成「列数极多且大量空格」的假表格
        // （实测 28 列几乎全空），当表格还原会把文字塞成一格一句、版面与页数双爆炸。
        // 真实表格很少超过 12 列、空单元格占比也不会这么高；命中守卫就按普通文本行
        // 处理（内容不丢，只是不还原成表格——注意直接 return，跳过下方删行逻辑）。
        const cells = cleaned.flat();
        const filled = cells.filter((c) => c).length / Math.max(1, cells.length);
        if (cleaned[0].length > 12 && filled < 0.4) return;
        // 标记在首行上：交给 linesToBlocks 按 y 位置插入，否则表格会被塞到页首
        inside[0].table = { rows: cleaned };
        // 从行集合中移除已归入表格的文本（保留首行作为表格锚点）
        for (const line of inside.slice(1)) {
            const idx = lines.indexOf(line);
            if (idx >= 0) lines.splice(idx, 1);
        }
        return;
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
