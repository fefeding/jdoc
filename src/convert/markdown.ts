/**
 * Markdown <-> 文档块模型（DocxBlock）互转
 *
 * 以 `markdownToBlocks` 把 Markdown 解析为库级统一 IR（`DocxBlock`），
 * 以 `blocksToMarkdown` 把 IR 渲染回 Markdown 文本。两条路径与 docx/html/pdf
 * 共用同一套块模型，因此任何格式都能经 `DocxBlock` 中转互相转换。
 *
 * 解析端直接用 marked 的词法器（lexer），保留行内样式（加粗/斜体/代码/链接）
 * 为 `DocxRun`，并对 mermaid / 数学公式走与 `markdownToHtml` 一致的扩展，
 * 保证 `md -> blocks -> md` 可无损往返。
 *
 * @module convert/markdown
 */

import { Marked, type Token, type Tokens } from 'marked';
import { markdownExtensions } from '../markdown/extensions';
import { getAsset, assetToDataUrl, toAssetTable, type DocAsset } from '../core/asset';
import { htmlToBlocks } from './html';
import type {
    DocxBlock,
    DocxCell,
    DocxImage,
    DocxListItem,
    DocxRun
} from '../docx/types';

/** Markdown 解析选项（与 markdown 模块对齐） */
export interface MarkdownToBlocksOptions {
    /** 是否启用 GFM（表格、删除线、任务列表、自动链接），默认 true */
    gfm?: boolean;
    /** 是否启用公式语法，默认 true */
    math?: boolean;
    /** 是否启用 mermaid 图表语法，默认 true */
    diagrams?: boolean;
}

/** blocksToMarkdown 选项 */
export interface BlocksToMarkdownOptions {
    /** 资源表，用于把图片块上的 assetId 解析为可内联的 data URL */
    assets?: DocAsset[] | Record<string, DocAsset> | { assets?: DocAsset[] };
    /** 是否把行内代码包成反引号，默认 true */
    inlineCode?: boolean;
}

// ---------------------------------------------------------------------------
// Markdown -> blocks
// ---------------------------------------------------------------------------

/** 实例缓存：按选项组合分别构建，避免第一次调用的选项“焊死”后续所有调用 */
const markedCache = new Map<string, Marked>();
function getMarked(options: MarkdownToBlocksOptions): Marked {
    const key = `${options.gfm !== false}|${options.math !== false}|${options.diagrams !== false}`;
    let instance = markedCache.get(key);
    if (!instance) {
        instance = new Marked({ gfm: options.gfm !== false, breaks: false });
        instance.use(
            markdownExtensions({ enableMath: options.math !== false, enableDiagrams: options.diagrams !== false })
        );
        markedCache.set(key, instance);
    }
    return instance;
}

/** 把若干行内 token 转成 run 序列 */
function inlineToRuns(tokens: Token[] | undefined): DocxRun[] {
    const runs: DocxRun[] = [];
    const pushStyled = (source: DocxRun[], style: Partial<DocxRun>) => {
        for (const r of source) runs.push({ ...r, ...style });
    };

    for (const t of tokens || []) {
        switch (t.type) {
            case 'text':
            case 'escape':
                runs.push({ text: (t as Tokens.Text | Tokens.Escape).text });
                break;
            case 'codespan':
                runs.push({ text: (t as Tokens.Codespan).text, code: true });
                break;
            case 'strong':
                pushStyled(inlineToRuns((t as Tokens.Strong).tokens), { bold: true });
                break;
            case 'em':
                pushStyled(inlineToRuns((t as Tokens.Em).tokens), { italic: true });
                break;
            case 'del':
                pushStyled(inlineToRuns((t as Tokens.Del).tokens), { strike: true });
                break;
            case 'link':
                pushStyled(inlineToRuns((t as Tokens.Link).tokens), { link: (t as Tokens.Link).href });
                break;
            case 'jdocMathInline':
                runs.push({ text: `$${(t as Token & { text?: string }).text ?? ''}$` });
                break;
            case 'br':
                runs.push({ text: '\n' });
                break;
            case 'image':
                runs.push({ text: `![${(t as Tokens.Image).text ?? ''}](${(t as Tokens.Image).href})` });
                break;
            case 'html':
                runs.push({ text: (t as Tokens.HTML).text });
                break;
            case 'space':
            default:
                break;
        }
    }
    return mergeRuns(runs);
}

/** 合并连续的同样式 run，避免碎片段 */
function mergeRuns(runs: DocxRun[]): DocxRun[] {
    const out: DocxRun[] = [];
    for (const r of runs) {
        const last = out[out.length - 1];
        if (last && sameStyle(last, r)) last.text += r.text;
        else out.push({ ...r });
    }
    return out;
}

/** 两个 run 是否样式相同（仅比较影响渲染的标志） */
function sameStyle(a: DocxRun, b: DocxRun): boolean {
    return (
        !!a.bold === !!b.bold &&
        !!a.italic === !!b.italic &&
        !!a.strike === !!b.strike &&
        !!a.code === !!b.code &&
        (a.link ?? '') === (b.link ?? '') &&
        (a.color ?? '') === (b.color ?? '') &&
        (a.font ?? '') === (b.font ?? '')
    );
}

/** 单元格 -> run 序列 */
function cellToRuns(cell: { tokens?: Token[] } | undefined): DocxRun[] {
    return inlineToRuns(cell?.tokens);
}

/** 把列表 token（含嵌套）转成 DocxList */
function listToBlock(token: Tokens.List): DocxBlock {
    const itemRuns = (item: Tokens.ListItem): DocxListItem => {
        const inline = (item.tokens || []).filter((t) => t.type !== 'list');
        const nested = (item.tokens || []).find((t) => t.type === 'list') as Tokens.List | undefined;
        const runs: DocxRun[] & { children?: DocxListItem[] } = inlineToRuns(inline);
        if (nested) runs.children = nested.items.map(itemRuns);
        return runs;
    };
    return { type: 'list', ordered: token.ordered, items: token.items.map(itemRuns) };
}

/** 取一个 token 流里的全部纯文本（用于引用块降级等） */
function tokensToText(tokens: Token[] | undefined): string {
    let text = '';
    for (const t of tokens || []) {
        if ('text' in t && typeof (t as { text?: string }).text === 'string') text += (t as { text: string }).text;
        else if ('tokens' in t && Array.isArray((t as { tokens?: Token[] }).tokens)) {
            text += tokensToText((t as { tokens: Token[] }).tokens);
        } else if ('items' in t && Array.isArray((t as { items?: Tokens.ListItem[] }).items)) {
            text += (t as { items: Tokens.ListItem[] }).items.map((i) => tokensToText(i.tokens)).join(' ');
        }
    }
    return text;
}

/**
 * 把 Markdown 解析为文档块模型
 * @param markdown - Markdown 文本（字符串）或 UTF-8 字节
 * @returns 块序列
 */
export function markdownToBlocks(markdown: string | Uint8Array, options: MarkdownToBlocksOptions = {}): DocxBlock[] {
    const text = typeof markdown === 'string' ? markdown : new TextDecoder('utf-8').decode(markdown);
    const tokens = getMarked(options).lexer(text);
    const blocks: DocxBlock[] = [];

    for (const token of tokens) {
        switch (token.type) {
            case 'heading': {
                const t = token as Tokens.Heading;
                const runs = t.tokens ? inlineToRuns(t.tokens) : undefined;
                blocks.push({
                    type: 'heading',
                    level: Math.min(6, Math.max(1, t.depth || 1)) as 1 | 2 | 3 | 4 | 5 | 6,
                    text: t.text,
                    runs
                });
                break;
            }
            case 'paragraph': {
                const t = token as Tokens.Paragraph;
                blocks.push({ type: 'paragraph', runs: inlineToRuns(t.tokens), text: t.text });
                break;
            }
            case 'list':
                blocks.push(listToBlock(token as Tokens.List));
                break;
            case 'table': {
                // marked 的表格 token 结构随版本略有差异，这里用宽松断言访问
                const t = token as unknown as {
                    header?: { tokens?: Token[] }[];
                    rows?: { cells?: { tokens?: Token[] }[]; tokens?: { tokens?: Token[] }[] }[];
                };
                const rows: DocxCell[][] = [];
                const header = (t.header || []).map((c) => [cellToRuns(c)] as DocxCell[]);
                rows.push(header as DocxCell[]);
                for (const row of t.rows || []) {
                    const cells = (row.cells || row.tokens || []) as { tokens?: Token[] }[];
                    rows.push(cells.map((c) => cellToRuns(c)) as DocxCell[]);
                }
                blocks.push({ type: 'table', header: true, rows });
                break;
            }
            case 'code': {
                const t = token as Tokens.Code;
                blocks.push({ type: 'code', text: t.text, language: t.lang || undefined });
                break;
            }
            case 'hr':
                blocks.push({ type: 'hr' });
                break;
            case 'blockquote': {
                const t = token as Tokens.Blockquote;
                blocks.push({ type: 'quote', runs: inlineToRuns(t.tokens), text: tokensToText(t.tokens) });
                break;
            }
            case 'jdocDiagramCode': {
                const t = token as Token & { lang?: string; text?: string };
                blocks.push({ type: 'code', text: t.text || '', language: t.lang || 'mermaid' });
                break;
            }
            case 'jdocMathBlock': {
                const t = token as Token & { text?: string };
                blocks.push({ type: 'code', text: t.text || '', language: 'math' });
                break;
            }
            case 'html': {
                // 原始 HTML 块：交给 htmlToBlocks 还原，保持 HTML 互转能力
                blocks.push(...htmlToBlocks((token as Tokens.HTML).text));
                break;
            }
            case 'space':
            default:
                break;
        }
    }
    return blocks;
}

// ---------------------------------------------------------------------------
// blocks -> Markdown
// ---------------------------------------------------------------------------

/** 行内 run 序列 -> Markdown 文本 */
function runsToMarkdown(runs: DocxRun[], inlineCode: boolean): string {
    let out = '';
    for (const r of runs) {
        if (!r.text) continue;
        let text = r.text;
        if (r.code && inlineCode) text = '`' + text + '`';
        else {
            if (r.bold) text = `**${text}**`;
            if (r.italic) text = `*${text}*`;
            if (r.strike) text = `~~${text}~~`;
        }
        if (r.link) text = `[${text}](${r.link})`;
        out += text;
    }
    return out;
}

/** 引用块 -> Markdown（每行补 `> ` 前缀，无引用内容时至少输出 `>`） */
function quoteToMarkdown(runs: DocxRun[] | undefined, text: string | undefined, inlineCode: boolean): string {
    const content = runs ? runsToMarkdown(runs, inlineCode) : text ?? '';
    if (!content.trim()) return '>';
    return content
        .split('\n')
        .map((line) => (line.trim() ? '> ' + line : '>'))
        .join('\n');
}

/** 单元格 -> 纯文本（用于 Markdown 表格） */
function cellToText(cell: DocxCell | undefined, inlineCode: boolean): string {
    if (cell == null) return '';
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return runsToMarkdown(cell, inlineCode);
    const detail = cell as { blocks?: DocxBlock[]; runs?: DocxRun[]; text?: string };
    if (detail.blocks) return detail.blocks.map((b) => blockToMarkdownLine(b, inlineCode)).join(' ');
    if (detail.runs) return runsToMarkdown(detail.runs, inlineCode);
    return detail.text ?? '';
}

/** 单个块转成单行 Markdown（不含结尾换行，供表格/引用等复用） */
function blockToMarkdownLine(block: DocxBlock, inlineCode: boolean): string {
    switch (block.type) {
        case 'heading':
            return `${'#'.repeat(block.level)} ${block.runs ? runsToMarkdown(block.runs, inlineCode) : block.text ?? ''}`;
        case 'paragraph':
            return block.runs ? runsToMarkdown(block.runs, inlineCode) : block.text ?? '';
        case 'code':
            return '```' + (block.language || '') + '\n' + block.text + '\n```';
        case 'image':
            return `![${block.alt ?? ''}](${imageSrc(block)})`;
        case 'hr':
            return '---';
        case 'pageBreak':
            return '<!-- pagebreak -->';
        case 'quote':
            // 引用块：保持 `> ` 前缀（内部换行按引用语法逐行补前缀）
            return quoteToMarkdown(block.runs, block.text, inlineCode);
        case 'textBox':
            return block.paragraphs.map((p) => blockToMarkdownLine(p, inlineCode)).join('\n');
        case 'list':
            return renderList(block, 0, inlineCode);
        case 'table':
            return renderTable(block, inlineCode);
        case 'toc':
            return '';
        default:
            return '';
    }
}

/** 列表渲染（支持嵌套，每层 2 空格缩进） */
function renderList(list: Extract<DocxBlock, { type: 'list' }>, depth: number, inlineCode: boolean): string {
    const indent = '  '.repeat(depth);
    const lines: string[] = [];
    list.items.forEach((item, index) => {
        // DocxListItem 允许 string 形态（等价于 [{text}]），统一归一为 runs
        const runs: DocxRun[] & { children?: DocxListItem[] } = typeof item === 'string' ? [{ text: item }] : (item as DocxRun[]);
        const children = runs.children;
        const prefix = list.ordered ? `${index + 1}. ` : '- ';
        lines.push(indent + prefix + runsToMarkdown(runs, inlineCode));
        if (children?.length) {
            const childList: Extract<DocxBlock, { type: 'list' }> = {
                type: 'list',
                ordered: list.ordered,
                items: children
            };
            lines.push(renderList(childList, depth + 1, inlineCode));
        }
    });
    return lines.join('\n');
}

/** 表格渲染（GFM 管道表） */
function renderTable(table: Extract<DocxBlock, { type: 'table' }>, inlineCode: boolean): string {
    if (!table.rows.length) return '';
    const cell = (c: DocxCell | undefined) => cellToText(c, inlineCode).replace(/\|/g, '\\|').replace(/\n/g, ' ');
    const header = table.rows[0].map(cell).join(' | ');
    const divider = table.rows[0].map(() => '---').join(' | ');
    const body = table.rows.slice(1).map((row) => row.map(cell).join(' | '));
    return ['| ' + header + ' |', '| ' + divider + ' |', ...body.map((r) => '| ' + r + ' |')].join('\n');
}

/** 当前渲染上下文的资源表（blocksToMarkdown 时设置） */
let assetLookup: Record<string, DocAsset> | undefined;

/** 图片源解析（assetId -> data URL） */
function imageSrc(image: DocxImage, assets?: Record<string, DocAsset>): string {
    const table = assets ?? assetLookup;
    if (typeof image.src === 'string') return image.src;
    if (image.assetId) {
        const asset = table ? getAsset(table, image.assetId) : undefined;
        if (asset) return assetToDataUrl(asset);
    }
    return '';
}

/** 把多种资源表形态归一为 id -> 资源 的索引表 */
function normalizeAssets(assets: BlocksToMarkdownOptions['assets']): Record<string, DocAsset> | undefined {
    if (!assets) return undefined;
    if (Array.isArray(assets)) return toAssetTable(assets);
    const explicit = (assets as { assets?: DocAsset[] }).assets;
    if (explicit) return toAssetTable(explicit);
    return assets as Record<string, DocAsset>;
}

/**
 * 把文档块模型渲染为 Markdown 文本
 * @param blocks - 块序列或完整文档
 * @param options - 资源表等
 */
export function blocksToMarkdown(
    input: DocxBlock[] | { blocks?: DocxBlock[] },
    options: BlocksToMarkdownOptions = {}
): string {
    const blocks = Array.isArray(input) ? input : input.blocks || [];
    const inlineCode = options.inlineCode !== false;
    const assetTable = normalizeAssets(options.assets);
    assetLookup = assetTable;
    try {
        const parts: string[] = [];

        for (const block of blocks) {
            const text = blockToMarkdownLine(block, inlineCode);
            // 相邻块之间一律空行：列表后的段落缺空行会被 CommonMark 当作
            // 列表项的 lazy continuation，相邻列表会合并、相邻段落会软换行合并，
            // 统一空行才能保证块边界在往返后不丢。
            if (parts.length) parts.push('');
            parts.push(text);
        }
        return parts.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    } finally {
        assetLookup = undefined;
    }
}
