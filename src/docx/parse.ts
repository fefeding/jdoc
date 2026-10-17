/**
 * DOCX 解析主入口
 *
 * 直接读取 OOXML 部件（document.xml / styles.xml / numbering.xml / rels），
 * 自研而非依赖 mammoth，是为了完整保留 run 级样式（颜色、字号、字体、高亮）
 * 与段落属性（对齐、缩进、行距），并保证解析结果可直接用于往返生成。
 *
 * @module docx/parse
 */

import JSZip from 'jszip';
import {
    attr,
    childElements,
    findAll,
    findFirst,
    firstChild,
    isElement,
    localName,
    parseXml,
    textOf,
    type XmlElement,
    type XmlNode
} from '../core/xml';
import { bytesToBase64, toUint8Array, type DocInput } from '../core/input';
import { JdocError } from '../core/errors';
import {
    headingLevelOf,
    mergeParagraphProps,
    mergeRunProps,
    quoteStyleOf,
    readParagraphProps,
    readRunProps,
    type StyleParagraphProps,
    type StyleRunProps
} from './style-props';
import { buildStylesIndex, emptyStylesIndex, type StylesIndex } from './styles';
import type {
    DocxBlock,
    DocxCell,
    DocxDocument,
    DocxImage,
    DocxListItem,
    DocxMetadata,
    DocxParagraph,
    DocxRun
} from './types';

/** EMU 转磅：1 磅 = 12700 EMU */
const EMU_PER_PT = 12700;

/** 关系表：rId -> { target, type, mode } */
interface RelEntry {
    target: string;
    type: string;
    /** External 关系的 target 是 URL */
    external: boolean;
}

/** 解析上下文 */
interface ParseContext {
    zip: JSZip;
    styles: StylesIndex;
    rels: Map<string, RelEntry>;
    /** numId -> 各级 numFmt（bullet / decimal / ...） */
    numbering: Map<string, string[]>;
    messages: string[];
    /** 图片部件路径 -> data URL */
    imageCache: Map<string, string>;
    includeImages: boolean;
}

/** 读取 ZIP 内文本部件，缺失返回空串 */
async function readText(zip: JSZip, path: string): Promise<string> {
    const file = zip.file(path);
    if (!file) return '';
    try {
        return await file.async('string');
    } catch {
        return '';
    }
}

/** 解析 word/_rels/document.xml.rels */
function parseRels(xml: string): Map<string, RelEntry> {
    const rels = new Map<string, RelEntry>();
    if (!xml) return rels;
    for (const node of parseXml(xml)) {
        if (!isElement(node)) continue;
        for (const rel of childElements(node, 'Relationship')) {
            const id = attr(rel, 'Id');
            const target = attr(rel, 'Target');
            if (!id || !target) continue;
            const mode = (attr(rel, 'TargetMode') || '').toLowerCase();
            rels.set(id, {
                target,
                type: attr(rel, 'Type') || '',
                external: mode === 'external'
            });
        }
    }
    return rels;
}

/** 解析 word/numbering.xml：numId -> 各层级 numFmt */
function parseNumbering(xml: string): Map<string, string[]> {
    const result = new Map<string, string[]>();
    if (!xml) return result;

    const roots = parseXml(xml);
    // abstractNumId -> numFmt[]
    const abstractFormats = new Map<string, string[]>();
    for (const abstractNum of findAll(roots, 'abstractNum')) {
        const abstractId = attr(abstractNum, 'abstractNumId');
        if (!abstractId) continue;
        const formats: string[] = [];
        for (const lvl of childElements(abstractNum, 'lvl')) {
            const ilvl = attr(lvl, 'ilvl');
            const numFmt = firstChild(lvl, 'numFmt');
            formats[Number(ilvl) || 0] = (numFmt && attr(numFmt, 'val')) || 'decimal';
        }
        abstractFormats.set(abstractId, formats);
    }

    for (const num of findAll(roots, 'num')) {
        const numId = attr(num, 'numId');
        const abstractNumId = firstChild(num, 'abstractNumId');
        const abstractId = abstractNumId && attr(abstractNumId, 'val');
        if (!numId) continue;
        result.set(numId, (abstractId && abstractFormats.get(abstractId)) || ['decimal']);
    }

    return result;
}

/** 定位部件的根元素（兼容是否带外层包裹，如 cp:coreProperties） */
function findPartRoot(roots: XmlNode[], preferred: string[]): XmlElement | undefined {
    for (const node of roots) {
        if (isElement(node) && preferred.includes(localName(node.tag))) return node;
    }
    const first = roots.find((node) => isElement(node));
    return first && isElement(first) ? first : undefined;
}

/** 读取 docProps/core.xml 元数据 */
function parseCoreProps(xml: string): DocxMetadata {
    const metadata: DocxMetadata = {};
    if (!xml) return metadata;
    // 元数据字段是根元素的子节点，必须先下钻到根，否则只会读到 root 本身
    const root = findPartRoot(parseXml(xml), ['coreProperties', 'CoreProperties', 'Properties']);
    if (!root) return metadata;

    for (const node of root.children) {
        if (!isElement(node)) continue;
        const name = localName(node.tag);
        const value = textOf(node).trim();
        if (!value) continue;
        if (name === 'title') metadata.title = value;
        else if (name === 'creator') metadata.creator = value;
        else if (name === 'subject') metadata.subject = value;
        else if (name === 'keywords') metadata.keywords = value;
        else if (name === 'description') metadata.description = value;
        else if (name === 'lastModifiedBy') metadata.lastModifiedBy = value;
        else if (name === 'Company') metadata.company = value;
    }
    if (metadata.creator && !metadata.author) metadata.author = metadata.creator;
    return metadata;
}

/** 读取 docProps/app.xml 元数据（主要取 Company） */
function parseAppProps(xml: string): Partial<DocxMetadata> {
    const result: Partial<DocxMetadata> = {};
    if (!xml) return result;
    const root = findPartRoot(parseXml(xml), ['Properties', 'ExtendedProperties']);
    if (!root) return result;

    for (const node of root.children) {
        if (!isElement(node)) continue;
        if (localName(node.tag) === 'Company') {
            const value = textOf(node).trim();
            if (value) result.company = value;
        }
    }
    return result;
}

/** 关系 target 解析为 ZIP 内绝对路径（rels Target 相对所在部件目录） */
function resolveTargetPath(baseDir: string, target: string): string {
    if (!target) return '';
    if (target.startsWith('/')) return target.slice(1);
    const stack = baseDir ? baseDir.split('/') : [];
    for (const seg of target.split('/')) {
        if (seg === '' || seg === '.') continue;
        if (seg === '..') stack.pop();
        else stack.push(seg);
    }
    return stack.join('/');
}

/** 读取图片部件并转为 data URL */
async function loadImageDataUrl(ctx: ParseContext, target: string): Promise<string | undefined> {
    if (!target) return undefined;
    const cached = ctx.imageCache.get(target);
    if (cached !== undefined) return cached || undefined;

    const file = ctx.zip.file(target);
    if (!file) {
        ctx.imageCache.set(target, '');
        ctx.messages.push(`图片部件缺失：${target}`);
        return undefined;
    }

    let dataUrl = '';
    try {
        const bytes = new Uint8Array(await file.async('uint8array'));
        const ext = (target.split('.').pop() || 'png').toLowerCase();
        const mime = ext === 'jpg' ? 'jpeg' : ext === 'svg' ? 'svg+xml' : ext;
        dataUrl = `data:image/${mime};base64,${bytesToBase64(bytes)}`;
    } catch (error) {
        ctx.messages.push(`图片读取失败：${target}`);
        dataUrl = '';
    }

    ctx.imageCache.set(target, dataUrl);
    return dataUrl || undefined;
}

/** run 级属性 -> DocxRun 字段 */
function applyRunProps(run: DocxRun, props: StyleRunProps): DocxRun {
    if (props.bold !== undefined) run.bold = props.bold;
    if (props.italic !== undefined) run.italic = props.italic;
    if (props.underline !== undefined) run.underline = props.underline;
    if (props.strike !== undefined) run.strike = props.strike;
    if (props.subscript !== undefined) run.subscript = props.subscript;
    if (props.superscript !== undefined) run.superscript = props.superscript;
    if (props.color) run.color = props.color;
    if (props.size !== undefined) run.size = props.size;
    if (props.font) run.font = props.font;
    if (props.highlight) run.highlight = props.highlight;
    return run;
}

/** 追加 run，自动合并与前一个完全相同样式的相邻 run */
function pushRun(runs: DocxRun[], run: DocxRun): void {
    if (!run.text) return;
    const prev = runs[runs.length - 1];
    if (prev && sameRunStyle(prev, run)) {
        prev.text += run.text;
        return;
    }
    runs.push(run);
}

/** 判断两个 run 的样式是否完全一致（用于合并相邻同样式 run） */
function sameRunStyle(a: DocxRun, b: DocxRun): boolean {
    const keys: (keyof DocxRun)[] = [
        'bold', 'italic', 'underline', 'strike', 'subscript', 'superscript',
        'color', 'size', 'font', 'highlight', 'link'
    ];
    return keys.every((key) => a[key] === b[key]);
}

/** 从 w:drawing / w:pict 中提取图片信息 */
async function extractImage(
    ctx: ParseContext,
    holder: XmlElement,
    altText?: string
): Promise<DocxImage | undefined> {
    // 关系 ID：DrawingML 用 r:embed，VML 老式图片用 r:id
    const blip = findFirst(holder, 'blip');
    let relId = blip ? attr(blip, 'embed') || attr(blip, 'link') : undefined;
    if (!relId) {
        const imagedata = findFirst(holder, 'imagedata');
        relId = imagedata ? attr(imagedata, 'id') || attr(imagedata, 'href') : undefined;
    }
    if (!relId) return undefined;

    const rel = ctx.rels.get(relId);
    if (!rel) {
        ctx.messages.push(`未找到图片关系：${relId}`);
        return undefined;
    }
    if (!ctx.includeImages) return undefined;

    const path = resolveTargetPath('word', rel.target);
    const src = await loadImageDataUrl(ctx, path);
    if (!src) return undefined;

    // wp:extent 以 EMU 记录显示尺寸
    const extent = findFirst(holder, 'extent');
    const cx = extent ? Number(attr(extent, 'cx')) : NaN;
    const cy = extent ? Number(attr(extent, 'cy')) : NaN;

    const image: DocxImage = { type: 'image', src };
    if (Number.isFinite(cx) && cx > 0) image.width = Math.round((cx / EMU_PER_PT) * 100) / 100;
    if (Number.isFinite(cy) && cy > 0) image.height = Math.round((cy / EMU_PER_PT) * 100) / 100;
    if (altText) image.alt = altText;
    return image;
}

/** 提取 wp:docPr 的 descr/title 作为 alt 文本 */
function readImageAlt(holder: XmlElement): string | undefined {
    const docPr = findFirst(holder, 'docPr');
    const descr = docPr ? attr(docPr, 'descr') : undefined;
    if (descr) return descr;
    const nvPr = findFirst(holder, 'cNvPr');
    const title = nvPr ? attr(nvPr, 'title') : undefined;
    return title || undefined;
}

/** 解析段落内所有 run / 超链接 / 图片 */
async function parseRuns(
    ctx: ParseContext,
    pEl: XmlElement,
    baseRun: StyleRunProps
): Promise<{ runs: DocxRun[]; images: DocxImage[] }> {
    const runs: DocxRun[] = [];
    const images: DocxImage[] = [];

    const walk = async (container: XmlElement, inherited: StyleRunProps, link?: string) => {
        for (const child of container.children) {
            if (!isElement(child)) continue;
            const name = localName(child.tag);

            if (name === 'r') {
                // run 级属性：run 自身 > 段落样式 > 文档默认
                const props = mergeRunProps(
                    mergeRunProps(ctx.styles.defaultRun, inherited),
                    readRunProps(firstChild(child, 'rPr'))
                );

                for (const piece of child.children) {
                    if (!isElement(piece)) continue;
                    const pieceName = localName(piece.tag);
                    if (pieceName === 't') {
                        const run: DocxRun = { text: textOf(piece) };
                        applyRunProps(run, props);
                        if (link) run.link = link;
                        pushRun(runs, run);
                    } else if (pieceName === 'br') {
                        const run: DocxRun = { text: '\n' };
                        applyRunProps(run, props);
                        if (link) run.link = link;
                        pushRun(runs, run);
                    } else if (pieceName === 'tab') {
                        const run: DocxRun = { text: '\t' };
                        applyRunProps(run, props);
                        if (link) run.link = link;
                        pushRun(runs, run);
                    } else if (pieceName === 'drawing' || pieceName === 'pict' || pieceName === 'object') {
                        const image = await extractImage(ctx, piece, readImageAlt(piece));
                        if (image) images.push(image);
                    }
                }
            } else if (name === 'hyperlink') {
                const relId = attr(child, 'id');
                const rel = relId ? ctx.rels.get(relId) : undefined;
                // w:anchor 为文档内锚点，Hyperlink 样式说明它是外链但缺关系时降级为锚点
                const href = rel ? rel.target : attr(child, 'anchor');
                await walk(child, inherited, href || undefined);
            } else if (name === 'smartTag' || name === 'sdt' || name === 'sdtContent' || name === 'ins' || name === 'bdo') {
                // 智能标签/内容控件/修订插入/双向覆盖：透明下钻
                await walk(child, inherited, link);
            }
        }
    };

    await walk(pEl, baseRun);
    return { runs, images };
}

/** 判断段落是否包含分页标记 */
function hasPageBreak(ctx: ParseContext, pEl: XmlElement): boolean {
    const pPr = firstChild(pEl, 'pPr');
    if (pPr && firstChild(pPr, 'pageBreakBefore')) return true;
    // w:br w:type="page"
    for (const br of findAll(pEl, 'br')) {
        if ((attr(br, 'type') || '').toLowerCase() === 'page') return true;
    }
    void ctx;
    return false;
}

/** 判断空段落是否带边框（水平线的形态） */
function hasParagraphBorder(pEl: XmlElement): boolean {
    const pPr = firstChild(pEl, 'pPr');
    if (!pPr) return false;
    const borders = firstChild(pPr, 'pBdr');
    if (!borders) return false;
    // 任意一侧有显式线型即视为边框段落
    return childElements(borders).some((side) => {
        const style = attr(side, 'val');
        return !!style && style.toLowerCase() !== 'none';
    });
}

/** 读取段落的有效属性（段落自身 pPr > 样式继承 > 文档默认） */
function resolveParagraphProps(ctx: ParseContext, pEl: XmlElement): {
    props: StyleParagraphProps;
    styleRun: StyleRunProps;
    styleName?: string;
    headingLevel?: 1 | 2 | 3 | 4 | 5 | 6;
} {
    const pPr = firstChild(pEl, 'pPr');
    const pStyleEl = pPr ? firstChild(pPr, 'pStyle') : undefined;
    const styleId = pStyleEl ? attr(pStyleEl, 'val') : undefined;
    const style = styleId ? ctx.styles.byId.get(styleId) : undefined;

    let props = mergeParagraphProps(ctx.styles.defaultParagraph, style ? style.paragraph : {});
    props = mergeParagraphProps(props, readParagraphProps(pPr));

    const styleRun = mergeRunProps(ctx.styles.defaultRun, style ? style.run : {});
    const headingLevel = headingLevelOf(style?.name);

    return { props, styleRun, styleName: style?.name, headingLevel };
}

/** 判断段落是否为空（无文本且无图片） */
function isEmptyParagraph(runs: DocxRun[], images: DocxImage[]): boolean {
    return !images.length && !runs.some((run) => run.text.trim().length > 0);
}

/** 段落内所有 run 合并为纯文本 */
function runsToText(runs: DocxRun[]): string {
    return runs.map((run) => run.text).join('');
}

/** 列表项：run 数组 + 子项 */
interface ListEntry {
    item: DocxListItem;
    /** 无序列表用 bullet，有序列表用 decimal */
    ordered: boolean;
}

/** 解析单个段落为块（可能产出多个块：图片独立成块） */
async function parseParagraph(ctx: ParseContext, pEl: XmlElement): Promise<DocxBlock[]> {
    const { props, styleRun, styleName, headingLevel } = resolveParagraphProps(ctx, pEl);
    const { runs, images } = await parseRuns(ctx, pEl, styleRun);
    const blocks: DocxBlock[] = [];

    for (const image of images) blocks.push(image);

    const empty = isEmptyParagraph(runs, []);

    // 显式分页：w:br[@w:type='page'] 或 w:pageBreakBefore
    if (hasPageBreak(ctx, pEl)) {
        blocks.push({ type: 'pageBreak' });
        if (empty && !images.length) return blocks;
    }

    // 空段落：仅在没有图片时输出占位段落（保留文档的段落结构）
    if (empty) {
        // 空段落带边框时按分隔线还原（生成端 hr 的形态）
        if (!images.length && hasParagraphBorder(pEl)) blocks.push({ type: 'hr' });
        else if (!images.length) blocks.push({ type: 'paragraph', text: '' });
        return blocks;
    }

    if (headingLevel !== undefined) {
        blocks.push({ type: 'heading', level: headingLevel, runs });
        return blocks;
    }

    if (quoteStyleOf(styleName)) {
        blocks.push({ type: 'quote', runs });
        return blocks;
    }

    // 段落级格式下沉到 run（对齐、缩进等属段落属性，这里只保留纯文本格式）
    const paragraph: DocxParagraph = { type: 'paragraph', runs };
    if (props.align) paragraph.align = props.align;
    if (props.indentLeft) paragraph.indentLeft = props.indentLeft;
    if (props.indentRight) paragraph.indentRight = props.indentRight;
    if (props.indentFirstLine) paragraph.indentFirstLine = props.indentFirstLine;
    if (props.spacingBefore) paragraph.spacingBefore = props.spacingBefore;
    if (props.spacingAfter) paragraph.spacingAfter = props.spacingAfter;
    if (props.lineSpacing && props.lineSpacing !== 1) paragraph.lineSpacing = props.lineSpacing;
    blocks.push(paragraph);

    return blocks;
}

/** 读取段落的列表信息 */
function readListInfo(
    ctx: ParseContext,
    pEl: XmlElement
): { ordered: boolean; level: number } | undefined {
    const pPr = firstChild(pEl, 'pPr');
    const numPr = pPr ? firstChild(pPr, 'numPr') : undefined;
    if (!numPr) return undefined;

    const numIdEl = firstChild(numPr, 'numId');
    const numId = numIdEl ? attr(numIdEl, 'val') : undefined;
    if (!numId || numId === '0') return undefined;

    const ilvlEl = firstChild(numPr, 'ilvl');
    const level = ilvlEl ? Number(attr(ilvlEl, 'val')) || 0 : 0;

    const formats = ctx.numbering.get(numId);
    const fmt = formats ? formats[level] || formats[0] : 'decimal';
    return { ordered: fmt !== 'bullet', level };
}

/** 解析表格单元格内容（单元格内可能含多个段落） */
async function parseCell(ctx: ParseContext, tcEl: XmlElement): Promise<DocxCell> {
    const runs: DocxRun[] = [];
    for (const child of tcEl.children) {
        if (!isElement(child) || localName(child.tag) !== 'p') continue;
        const { styleRun } = resolveParagraphProps(ctx, child);
        const parsed = await parseRuns(ctx, child, styleRun);
        runs.push(...parsed.runs);
    }
    return runs;
}

/** 解析表格 */
async function parseTable(ctx: ParseContext, tblEl: XmlElement): Promise<DocxBlock> {
    const rows: DocxCell[][] = [];

    for (const child of tblEl.children) {
        if (!isElement(child) || localName(child.tag) !== 'tr') continue;
        const cells: DocxCell[] = [];
        for (const cellChild of child.children) {
            if (!isElement(cellChild) || localName(cellChild.tag) !== 'tc') continue;
            cells.push(await parseCell(ctx, cellChild));
        }
        if (cells.length) rows.push(cells);
    }

    return { type: 'table', rows, header: rows.length > 0 };
}

/**
 * 遍历 w:body，把相邻的同类型列表段落合并为一个 list 块
 */
async function parseBody(ctx: ParseContext, body: XmlElement): Promise<DocxBlock[]> {
    const blocks: DocxBlock[] = [];
    // 当前正在累积的列表（用于合并连续列表项）
    let activeList: { type: 'list'; ordered: boolean; items: DocxListItem[] } | undefined;
    let activeLevel = 0;

    const flushList = () => {
        if (activeList && activeList.items.length) blocks.push(activeList);
        activeList = undefined;
        activeLevel = 0;
    };

    for (const child of body.children) {
        if (!isElement(child)) continue;
        const name = localName(child.tag);

        if (name === 'p') {
            const listInfo = readListInfo(ctx, child);
            const { styleRun } = resolveParagraphProps(ctx, child);
            const { runs, images } = await parseRuns(ctx, child, styleRun);

            // 列表段落：图片不并入列表项，单独成块并保持列表继续
            for (const image of images) {
                flushList();
                blocks.push(image);
            }

            if (listInfo) {
                const isSameList =
                    activeList && activeList.ordered === listInfo.ordered && activeLevel === listInfo.level;

                if (!isSameList) {
                    flushList();
                    activeList = { type: 'list', ordered: listInfo.ordered, items: [] };
                    activeLevel = listInfo.level;
                }

                const text = runsToText(runs);
                // 多级列表：把子级列表作为当前项的 children
                if (listInfo.level > activeLevel && activeList.items.length) {
                    const last = activeList.items[activeList.items.length - 1];
                    if (typeof last !== 'string' && Array.isArray(last)) {
                        (last as DocxListItem & { children?: DocxListItem[] }).children =
                            (last as DocxListItem & { children?: DocxListItem[] }).children || [];
                        (last as DocxListItem & { children?: DocxListItem[] }).children!.push(text);
                    }
                } else {
                    activeList.items.push(runs.length > 1 ? runs : text);
                }
                continue;
            }

            flushList();
            blocks.push(...(await parseParagraph(ctx, child)));
            continue;
        }

        if (name === 'tbl') {
            flushList();
            blocks.push(await parseTable(ctx, child));
            continue;
        }

        // sectPr / bookmark / proofErr 等其他节点忽略
    }

    flushList();
    return blocks;
}

/** 解析结果 */
export interface DocxParseOutput {
    document: DocxDocument;
    /** 解析过程中的提示（缺部件、图片丢失等） */
    messages: string[];
    /** 全文纯文本 */
    text: string;
}

/** 从块序列提取纯文本（表格按行用制表符分隔） */
export function blocksToText(blocks: DocxBlock[]): string {
    const lines: string[] = [];
    for (const block of blocks) {
        switch (block.type) {
            case 'heading':
            case 'paragraph':
            case 'quote':
                lines.push(typeof block.text === 'string' ? block.text : runsToText(block.runs || []));
                break;
            case 'code':
                lines.push(block.text ?? '');
                break;
            case 'list':
                for (const item of block.items) {
                    lines.push(typeof item === 'string' ? item : runsToText(item));
                }
                break;
            case 'table':
                for (const row of block.rows) {
                    lines.push(
                        row
                            .map((cell) => (typeof cell === 'string' ? cell : runsToText(cell)))
                            .join('\t')
                    );
                }
                break;
            case 'image':
                if (block.caption) lines.push(block.caption);
                break;
            default:
                break;
        }
    }
    return lines.join('\n');
}

/**
 * 解析 DOCX 为结构化文档（内部函数）
 * @param input - 文件二进制
 * @param includeImages - 是否内联图片为 data URL
 */
export async function parseDocxInternal(
    input: DocInput,
    includeImages = true
): Promise<DocxParseOutput> {
    const bytes = await toUint8Array(input);
    if (bytes.byteLength < 4) {
        throw new JdocError('docx 数据无效：文件过小', 'docx');
    }

    const zip = await JSZip.loadAsync(bytes);
    const messages: string[] = [];

    const [documentXml, stylesXml, numberingXml, relsXml, coreXml, appXml] = await Promise.all([
        readText(zip, 'word/document.xml'),
        readText(zip, 'word/styles.xml'),
        readText(zip, 'word/numbering.xml'),
        readText(zip, 'word/_rels/document.xml.rels'),
        readText(zip, 'docProps/core.xml'),
        readText(zip, 'docProps/app.xml')
    ]);

    if (!documentXml) {
        throw new JdocError('docx 结构异常：缺少 word/document.xml', 'docx');
    }

    const ctx: ParseContext = {
        zip,
        styles: stylesXml ? buildStylesIndex(stylesXml) : emptyStylesIndex(),
        rels: parseRels(relsXml),
        numbering: parseNumbering(numberingXml),
        messages,
        imageCache: new Map(),
        includeImages
    };

    const roots = parseXml(documentXml);
    const body = findFirst(roots, 'body');
    const blocks = body ? await parseBody(ctx, body) : [];

    const metadata = { ...parseCoreProps(coreXml), ...parseAppProps(appXml) };
    const document: DocxDocument = {
        metadata: Object.keys(metadata).length ? metadata : undefined,
        blocks
    };

    return { document, messages, text: blocksToText(blocks) };
}
