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
    attrNumber,
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
import {
    createAsset,
    createAssetId,
    type DocAsset,
    type DocAssetMode
} from '../core/asset';
import { JdocError } from '../core/errors';
import {
    headingLevelOf,
    mergeParagraphProps,
    mergeRunProps,
    quoteStyleOf,
    readParagraphProps,
    readRunProps,
    type StyleParagraphProps,
    type StyleRunProps,
    type ThemeFonts
} from './style-props';
import {
    buildStylesIndex,
    emptyStylesIndex,
    parseThemeColors,
    parseThemeFonts,
    resolveDrawingColor,
    type StylesIndex,
    type ThemeColors
} from './styles';
import type {
    DocxBlock,
    DocxCell,
    DocxCellBorder,
    DocxCellDetail,
    DocxDocument,
    DocxImage,
    DocxListItem,
    DocxMetadata,
    DocxParagraph,
    DocxRun,
    DocxSection,
    DocxTable,
    DocxTextBox
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
    /** 二进制资源处理方式 */
    assetMode: DocAssetMode;
    /** 已提取的资源（按部件路径去重，同一张图只存一份） */
    assets: DocAsset[];
    /** 部件路径 -> 资源 ID */
    assetIds: Map<string, string>;
    /** 主题字体（theme1.xml），用于还原 asciiTheme/eastAsiaTheme */
    themeFonts?: ThemeFonts;
    /** 主题色板（theme1.xml），用于还原 schemeClr 引用 */
    themeColors: ThemeColors;
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

    // wp:extent 以 EMU 记录显示尺寸
    const extent = findFirst(holder, 'extent');
    const cx = extent ? Number(attr(extent, 'cx')) : NaN;
    const cy = extent ? Number(attr(extent, 'cy')) : NaN;
    const width = Number.isFinite(cx) && cx > 0 ? Math.round((cx / EMU_PER_PT) * 100) / 100 : undefined;
    const height = Number.isFinite(cy) && cy > 0 ? Math.round((cy / EMU_PER_PT) * 100) / 100 : undefined;

    const extract = ctx.assetMode === 'extract' || ctx.assetMode === 'both';
    const image: DocxImage = { type: 'image', src: '' };

    // 提取模式：字节进资源表，块上留引用（同一张图多处引用只存一份）
    if (extract) {
        const assetId = await registerAsset(ctx, path, { width, height, alt: altText });
        if (!assetId) return undefined;
        image.assetId = assetId;
        image.src = { assetId };
        if (ctx.assetMode === 'extract') return finishImage(image, width, height, altText);
    }

    // 内联模式（inline，或 both 的显示侧）：data URL 直接可用
    const src = await loadImageDataUrl(ctx, path);
    if (!src) return undefined;
    image.src = src;
    return finishImage(image, width, height, altText);
}

/** 统一补齐图片块的尺寸与 alt */
function finishImage(
    image: DocxImage,
    width: number | undefined,
    height: number | undefined,
    altText: string | undefined
): DocxImage {
    if (width != null) image.width = width;
    if (height != null) image.height = height;
    if (altText) image.alt = altText;
    return image;
}

/**
 * 把 ZIP 内的图片部件注册为文档资源，返回资源 ID
 *
 * 按部件路径去重：同一张图在文档里出现多次时共享一个资源 ID 与一份字节。
 */
async function registerAsset(
    ctx: ParseContext,
    path: string,
    info: { width?: number; height?: number; alt?: string }
): Promise<string | undefined> {
    const cached = ctx.assetIds.get(path);
    if (cached) return cached;

    const file = ctx.zip.file(path);
    if (!file) {
        ctx.messages.push(`图片部件缺失：${path}`);
        return undefined;
    }

    try {
        const data = new Uint8Array(await file.async('uint8array'));
        const id = createAssetId('docx', ctx.assets.length + 1);
        const asset = createAsset({
            id,
            data,
            part: path,
            anchor: { scope: 'docx', kind: 'block', index: -1 },
            width: info.width,
            height: info.height,
            alt: info.alt
        });
        ctx.assets.push(asset);
        ctx.assetIds.set(path, id);
        return id;
    } catch (error) {
        ctx.messages.push(`图片读取失败：${path}`);
        return undefined;
    }
}

/**
 * 回填资源的位置锚点：遍历块序列，为每个被引用的资源标出它属于哪个块 / 单元格
 *
 * 这样「资源 ↔ 文档位置」是双向可查的：既能由块找到图片字节，
 * 也能由图片反查它在文档中的落点（例如「第 12 块 / Sheet1!B3」）。
 */
function assignAssetAnchors(blocks: DocxBlock[], assets: DocAsset[]): void {
    if (!assets.length) return;
    const index = new Map(assets.map((asset) => [asset.id, asset]));

    const walk = (list: DocxBlock[], prefix: string) => {
        list.forEach((block, i) => {
            const path = `${prefix}[${i}]`;
            if (block.type === 'image' && block.assetId) {
                const asset = index.get(block.assetId);
                if (asset) {
                    asset.anchor = { scope: 'docx', kind: 'block', index: i, path, label: `第 ${i + 1} 个块` };
                    if (block.width != null) asset.width ??= block.width;
                    if (block.height != null) asset.height ??= block.height;
                }
            }
            if (block.type === 'table') {
                const rows = block.rows || [];
                rows.forEach((row, r) => {
                    row.forEach((cell, c) => {
                        const detail = (cell as { detail?: { blocks?: DocxBlock[] } })?.detail;
                        if (detail?.blocks) walk(detail.blocks, `${path}.rows[${r}].cells[${c}]`);
                    });
                });
            }
        });
    };
    walk(blocks, 'blocks');
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

/** 解析段落内所有 run / 超链接 / 图片 / 文本框 */
async function parseRuns(
    ctx: ParseContext,
    pEl: XmlElement,
    baseRun: StyleRunProps
): Promise<{ runs: DocxRun[]; images: DocxImage[]; textBoxes: DocxTextBox[] }> {
    const runs: DocxRun[] = [];
    const images: DocxImage[] = [];
    const textBoxes: DocxTextBox[] = [];

    const walk = async (container: XmlElement, inherited: StyleRunProps, link?: string) => {
        for (const child of container.children) {
            if (!isElement(child)) continue;
            const name = localName(child.tag);

            if (name === 'r') {
                // run 级属性：run 自身 > 段落样式 > 文档默认
                const props = mergeRunProps(
                    mergeRunProps(ctx.styles.defaultRun, inherited),
                    readRunProps(firstChild(child, 'rPr'), ctx.themeFonts)
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
                        // DrawingML/VML 文本框（代码片段、标注框）：内容按段落提取
                        const textBox = await extractTextBox(ctx, piece);
                        if (textBox) textBoxes.push(textBox);
                    } else if (pieceName === 'AlternateContent') {
                        // mc:Choice 优先（新格式），缺失时回退 mc:Fallback，避免同一内容渲染两次
                        const branch = firstChild(piece, 'Choice') || firstChild(piece, 'Fallback');
                        if (branch) {
                            const holder = firstChild(branch, 'drawing') || firstChild(branch, 'pict');
                            if (holder) {
                                const image = await extractImage(ctx, holder, readImageAlt(holder));
                                if (image) images.push(image);
                                const textBox = await extractTextBox(ctx, holder);
                                if (textBox) textBoxes.push(textBox);
                            }
                        }
                    }
                }
            } else if (name === 'hyperlink') {
                const relId = attr(child, 'id');
                const rel = relId ? ctx.rels.get(relId) : undefined;
                // w:anchor 为文档内锚点，Hyperlink 样式说明它是外链但缺关系时降级为锚点
                const href = rel ? rel.target : attr(child, 'anchor');
                await walk(child, inherited, href || undefined);
            } else if (name === 'AlternateContent') {
                // mc:Choice 优先（新格式），缺失时回退 mc:Fallback，避免同一内容渲染两次
                const branch = firstChild(child, 'Choice') || firstChild(child, 'Fallback');
                if (branch) await walk(branch, inherited, link);
            } else if (name === 'smartTag' || name === 'sdt' || name === 'sdtContent' || name === 'ins' || name === 'bdo') {
                // 智能标签/内容控件/修订插入/双向覆盖：透明下钻
                await walk(child, inherited, link);
            }
        }
    };

    await walk(pEl, baseRun);
    return { runs, images, textBoxes };
}

/** 从 w:drawing / w:pict 中提取文本框内容（w:txbxContent 内的段落） */
async function extractTextBox(ctx: ParseContext, holder: XmlElement): Promise<DocxTextBox | undefined> {
    const txbx = findFirst(holder, 'txbxContent');
    if (!txbx) return undefined;

    const paragraphs: DocxBlock[] = [];
    for (const child of txbx.children) {
        if (!isElement(child) || localName(child.tag) !== 'p') continue;
        paragraphs.push(...(await parseParagraph(ctx, child)));
    }
    if (!paragraphs.length) return undefined;

    const box: DocxTextBox = { type: 'textBox', paragraphs };

    // 填充色：wps:spPr 的直接子 a:solidFill（srgbClr 或 schemeClr 主题色），
    // 或 VML 的 fillcolor（形如 "#222A35 [1615]"）
    const spPr = findFirst(holder, 'spPr');
    const fillEl = spPr ? firstChild(spPr, 'solidFill') : undefined;
    const fillColorEl = fillEl
        ? firstChild(fillEl, 'srgbClr') || firstChild(fillEl, 'schemeClr')
        : undefined;
    const fill = resolveDrawingColor(fillColorEl, ctx.themeColors);
    if (fill) box.fill = fill;
    const vmlShape = findFirst(holder, 'shape');
    const fillcolor = vmlShape ? attr(vmlShape, 'fillcolor') : undefined;
    if (!box.fill && fillcolor) {
        const hex = fillcolor.trim().split(/\s+/)[0].replace(/^#/, '');
        if (/^[0-9a-fA-F]{6}$/.test(hex)) box.fill = `#${hex.toUpperCase()}`;
    }
    // 边框色：wps:spPr 的 a:ln > a:solidFill，或 VML strokecolor
    const lnEl = spPr ? firstChild(spPr, 'ln') : undefined;
    const lnFill = lnEl ? firstChild(lnEl, 'solidFill') : undefined;
    const stroke = resolveDrawingColor(
        lnFill ? firstChild(lnFill, 'srgbClr') || firstChild(lnFill, 'schemeClr') : undefined,
        ctx.themeColors
    );
    if (stroke) box.stroke = stroke;
    const strokecolor = vmlShape ? attr(vmlShape, 'strokecolor') : undefined;
    if (!box.stroke && strokecolor && strokecolor.toLowerCase() !== 'none') {
        const hex = strokecolor.trim().split(/\s+/)[0].replace(/^#/, '');
        if (/^[0-9a-fA-F]{6}$/.test(hex)) box.stroke = `#${hex.toUpperCase()}`;
    }

    const extent = findFirst(holder, 'extent');
    const cx = extent ? Number(attr(extent, 'cx')) : NaN;
    const cy = extent ? Number(attr(extent, 'cy')) : NaN;
    if (Number.isFinite(cx) && cx > 0) box.width = Math.round((cx / EMU_PER_PT) * 100) / 100;
    if (Number.isFinite(cy) && cy > 0) box.height = Math.round((cy / EMU_PER_PT) * 100) / 100;

    return box;
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
    // 未标 pStyle 的段落使用文档默认段落样式（Word 的 Normal：两端对齐、默认字号等）
    const styleId = pStyleEl
        ? attr(pStyleEl, 'val')
        : ctx.styles.defaultParagraphStyleId;
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

/** 段落属性 -> 块上的可选布局字段（heading/quote 也需要间距对齐） */
function copyParagraphProps(props: StyleParagraphProps): Partial<DocxParagraph> {
    const out: Partial<DocxParagraph> = {};
    if (props.align) out.align = props.align;
    if (props.indentLeft) out.indentLeft = props.indentLeft;
    if (props.indentRight) out.indentRight = props.indentRight;
    if (props.indentFirstLine) out.indentFirstLine = props.indentFirstLine;
    if (props.spacingBefore) out.spacingBefore = props.spacingBefore;
    if (props.spacingAfter) out.spacingAfter = props.spacingAfter;
    if (props.lineSpacing && props.lineSpacing !== 1) out.lineSpacing = props.lineSpacing;
    return out;
}

/** 解析单个段落为块（可能产出多个块：图片/文本框独立成块） */
async function parseParagraph(ctx: ParseContext, pEl: XmlElement): Promise<DocxBlock[]> {
    const { props, styleRun, styleName, headingLevel } = resolveParagraphProps(ctx, pEl);
    const { runs, images, textBoxes } = await parseRuns(ctx, pEl, styleRun);
    const blocks: DocxBlock[] = [];

    for (const image of images) blocks.push(image);
    for (const box of textBoxes) blocks.push(box);

    const empty = isEmptyParagraph(runs, []);

    // 显式分页：w:br[@w:type='page'] 或 w:pageBreakBefore
    if (hasPageBreak(ctx, pEl)) {
        blocks.push({ type: 'pageBreak' });
        if (empty && !images.length && !textBoxes.length) return blocks;
    }

    // 空段落：仅在没有图片/文本框时输出占位段落（保留文档的段落结构）
    if (empty) {
        if (!images.length && !textBoxes.length) {
            // 空段落带边框时按分隔线还原（生成端 hr 的形态）
            if (hasParagraphBorder(pEl)) blocks.push({ type: 'hr' });
            else blocks.push({ type: 'paragraph', text: '' });
        }
        return blocks;
    }

    if (headingLevel !== undefined) {
        blocks.push({ type: 'heading', level: headingLevel, runs, ...copyParagraphProps(props) });
        return blocks;
    }

    if (quoteStyleOf(styleName)) {
        blocks.push({ type: 'quote', runs, ...copyParagraphProps(props) });
        return blocks;
    }

    // 段落级格式（对齐、缩进、间距）附在段落块上
    const paragraph: DocxParagraph = { type: 'paragraph', runs, ...copyParagraphProps(props) };
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

/** 读取单元格边框单边（w:tcBorders 的 top/left/bottom/right） */
function readCellBorder(sideEl: XmlElement | undefined): DocxCellBorder | undefined {
    if (!sideEl) return undefined;
    const style = (attr(sideEl, 'val') || '').toLowerCase();
    if (!style || style === 'none') return undefined;
    if (style === 'nil') return { style: 'nil' };
    const color = attr(sideEl, 'color');
    const size = attrNumber(sideEl, 'sz', 0);
    const border: DocxCellBorder = { style };
    if (color && color.toLowerCase() !== 'auto') border.color = `#${color.replace(/^#/, '')}`;
    // w:sz 单位是 1/8 磅
    if (size > 0) border.size = Math.round((size / 8) * 100) / 100;
    return border;
}

/** 解析表格单元格（含 w:tcPr 属性与多段落内容） */
async function parseCell(
    ctx: ParseContext,
    tcEl: XmlElement
): Promise<{ cell: DocxCell; span: number; vMerge?: 'restart' | 'continue' }> {
    const tcPr = firstChild(tcEl, 'tcPr');

    // 跨列（w:gridSpan）
    const spanEl = tcPr ? firstChild(tcPr, 'gridSpan') : undefined;
    const span = spanEl ? Math.max(1, attrNumber(spanEl, 'val', 1)) : 1;

    // 跨行（w:vMerge；无 val 属性默认为 continue）
    const vMergeEl = tcPr ? firstChild(tcPr, 'vMerge') : undefined;
    const vMerge = vMergeEl ? ((attr(vMergeEl, 'val') || 'continue') as 'restart' | 'continue') : undefined;

    // 宽度（w:tcW，dxa -> 磅）
    const widthEl = tcPr ? firstChild(tcPr, 'tcW') : undefined;
    let width: number | undefined;
    if (widthEl && (attr(widthEl, 'type') || 'dxa').toLowerCase() === 'dxa') {
        const w = attrNumber(widthEl, 'w', 0);
        if (w > 0) width = Math.round((w / 20) * 100) / 100;
    }

    // 底纹（w:shd fill）
    const shdEl = tcPr ? firstChild(tcPr, 'shd') : undefined;
    let fill: string | undefined;
    if (shdEl) {
        const shade = attr(shdEl, 'val');
        const color = attr(shdEl, 'fill');
        if (color && color.toLowerCase() !== 'auto' && (!shade || shade.toLowerCase() !== 'nil')) {
            fill = `#${color.replace(/^#/, '')}`;
        }
    }

    // 垂直对齐
    const vAlignEl = tcPr ? firstChild(tcPr, 'vAlign') : undefined;
    const vAlignRaw = vAlignEl ? (attr(vAlignEl, 'val') || '').toLowerCase() : '';
    const verticalAlign =
        vAlignRaw === 'center' ? 'middle' : vAlignRaw === 'bottom' ? 'bottom' : vAlignRaw === 'top' ? 'top' : undefined;

    // 边框
    const bordersEl = tcPr ? firstChild(tcPr, 'tcBorders') : undefined;
    let borders: DocxCellDetail['borders'] | undefined;
    if (bordersEl) {
        const sides = ['top', 'left', 'bottom', 'right'] as const;
        const read: Record<string, DocxCellBorder | undefined> = {};
        for (const side of sides) {
            read[side] = readCellBorder(firstChild(bordersEl, side));
        }
        if (Object.values(read).some(Boolean)) {
            borders = {
                top: read.top,
                left: read.left,
                bottom: read.bottom,
                right: read.right
            };
        }
    }

    // 内容：逐段落解析，保留段落级样式（对齐/缩进）
    const blocks: DocxBlock[] = [];
    for (const child of tcEl.children) {
        if (!isElement(child) || localName(child.tag) !== 'p') continue;
        blocks.push(...(await parseParagraph(ctx, child)));
    }

    const hasProps = span > 1 || !!vMerge || width !== undefined || !!fill || !!verticalAlign || !!borders;
    let cell: DocxCell;
    if (!hasProps && blocks.length === 1 && blocks[0].type === 'paragraph') {
        // 简写形态：单段落无属性时保持旧的富文本/字符串形态
        const para = blocks[0] as DocxParagraph;
        cell = para.runs && para.runs.length ? para.runs : para.text || '';
    } else if (!hasProps && blocks.length === 0) {
        cell = '';
    } else {
        const detail: DocxCellDetail = { type: 'cell' };
        // 内容尽量精简：全部是段落时存 blocks
        if (blocks.length) detail.blocks = blocks;
        if (span > 1) detail.colspan = span;
        if (width !== undefined) detail.width = width;
        if (fill) detail.fill = fill;
        if (verticalAlign) detail.verticalAlign = verticalAlign;
        if (borders) detail.borders = borders;
        cell = detail;
    }

    return { cell, span, vMerge };
}

/** 解析表格（tblGrid 列宽 + gridSpan/vMerge 合并 + 单元格属性） */
async function parseTable(ctx: ParseContext, tblEl: XmlElement): Promise<DocxBlock> {
    // 列宽（w:tblGrid/w:gridCol，dxa -> 磅）
    const gridEl = findFirst(tblEl, 'tblGrid');
    const columnWidths = gridEl
        ? childElements(gridEl, 'gridCol')
              .map((col) => Math.round((attrNumber(col, 'w', 0) / 20) * 100) / 100)
              .filter((w) => w > 0)
        : undefined;

    const rows: DocxCell[][] = [];
    // vMerge 向下合并的来源单元格：列下标 -> { cell, span }
    const rowSpanSources = new Map<number, { cell: DocxCellDetail; span: number }>();
    let firstRowIsHeader = false;

    for (const child of tblEl.children) {
        if (!isElement(child) || localName(child.tag) !== 'tr') continue;

        // 行属性：表头行（w:tblHeader，跨页重复）
        const trPr = firstChild(child, 'trPr');
        const isHeaderRow = rows.length === 0 && !!trPr && !!firstChild(trPr, 'tblHeader');
        if (isHeaderRow) firstRowIsHeader = true;

        const cells: DocxCell[] = [];
        let col = 0;
        for (const cellChild of child.children) {
            if (!isElement(cellChild) || localName(cellChild.tag) !== 'tc') continue;
            const { cell, span, vMerge } = await parseCell(ctx, cellChild);

            if (vMerge === 'continue') {
                // 延续行：不产新单元格，累加来源单元格的 rowspan
                const source = rowSpanSources.get(col);
                if (source && source.cell.rowspan !== undefined) source.cell.rowspan += 1;
                col += span;
                continue;
            }

            const detail = cell !== null && typeof cell === 'object' && !Array.isArray(cell) && cell.type === 'cell'
                ? cell
                : undefined;
            if (vMerge === 'restart' && detail) {
                detail.rowspan = 1;
                for (let i = 0; i < span; i++) rowSpanSources.set(col + i, { cell: detail, span });
            } else if (vMerge === 'restart') {
                // 简写形态无法承载 rowspan，忽略
            } else {
                for (let i = 0; i < span; i++) rowSpanSources.delete(col + i);
            }

            cells.push(cell);
            col += span;
        }

        // 整行都是 vMerge 延续时不产出行（HTML 由 rowspan 自动占位）
        if (cells.length) rows.push(cells);
    }

    const table: DocxTable = { type: 'table', rows, header: firstRowIsHeader };
    if (columnWidths && columnWidths.length) table.columnWidths = columnWidths;
    return table;
}

/** 解析 w:sectPr 页面设置（取 body 级最后一个 sectPr，即文档末尾的节属性） */
function parseSection(body: XmlElement): DocxSection | undefined {
    let sectPr: XmlElement | undefined;
    for (const child of body.children) {
        if (isElement(child) && localName(child.tag) === 'sectPr') sectPr = child;
    }
    if (!sectPr) return undefined;

    const pgSz = firstChild(sectPr, 'pgSz');
    const pgMar = firstChild(sectPr, 'pgMar');
    if (!pgSz) return undefined;

    const width = attrNumber(pgSz, 'w', 0) / 20;
    const height = attrNumber(pgSz, 'h', 0) / 20;
    if (width <= 0 || height <= 0) return undefined;

    const orientation = (attr(pgSz, 'orient') || '').toLowerCase() === 'landscape' ? 'landscape' : 'portrait';
    // pgMar 为自闭合元素，边距直接是属性（w:top/w:right/w:bottom/w:left，单位 twips）
    const mar = (name: string) => {
        const twips = pgMar ? attrNumber(pgMar, name, 1440) : 1440;
        return Math.round((twips / 20) * 100) / 100;
    };

    return {
        width: Math.round(width * 100) / 100,
        height: Math.round(height * 100) / 100,
        orientation,
        marginTop: mar('top'),
        marginRight: mar('right'),
        marginBottom: mar('bottom'),
        marginLeft: mar('left')
    };
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

            // 非列表段落直接走 parseParagraph（避免 parseRuns 跑两遍导致图片重复产出）
            if (!listInfo) {
                flushList();
                blocks.push(...(await parseParagraph(ctx, child)));
                continue;
            }

            const { styleRun } = resolveParagraphProps(ctx, child);
            const { runs, images, textBoxes } = await parseRuns(ctx, child, styleRun);

            // 列表段落：图片/文本框不并入列表项，单独成块并保持列表继续
            for (const image of images) {
                flushList();
                blocks.push(image);
            }
            for (const box of textBoxes) {
                flushList();
                blocks.push(box);
            }

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

/** 单元格 -> 纯文本 */
function cellToText(cell: DocxCell): string {
    if (typeof cell === 'string') return cell;
    if (Array.isArray(cell)) return runsToText(cell);
    // 对象形态：优先 blocks，其次 runs/text
    if (cell.blocks && cell.blocks.length) return blocksToText(cell.blocks);
    if (cell.runs) return runsToText(cell.runs);
    return cell.text ?? '';
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
                    lines.push(row.map((cell) => cellToText(cell)).join('\t'));
                }
                break;
            case 'image':
                if (block.caption) lines.push(block.caption);
                break;
            case 'textBox':
                lines.push(blocksToText(block.paragraphs));
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
    includeImages = true,
    assets: DocAssetMode = 'inline'
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

    // 主题字体：优先按关系表定位 theme 部件，退回约定路径
    const rels = parseRels(relsXml);
    let themePath = 'word/theme/theme1.xml';
    for (const rel of rels.values()) {
        if (rel.type.endsWith('/theme')) {
            themePath = resolveTargetPath('word', rel.target);
            break;
        }
    }
    const themeXml = await readText(zip, themePath);
    const themeFonts = parseThemeFonts(themeXml);
    const themeColors = parseThemeColors(themeXml);

    const ctx: ParseContext = {
        zip,
        styles: stylesXml ? buildStylesIndex(stylesXml, themeFonts) : emptyStylesIndex(),
        rels,
        numbering: parseNumbering(numberingXml),
        messages,
        imageCache: new Map(),
        includeImages,
        assetMode: assets,
        assets: [],
        assetIds: new Map(),
        themeFonts,
        themeColors
    };

    const roots = parseXml(documentXml);
    const body = findFirst(roots, 'body');
    const blocks = body ? await parseBody(ctx, body) : [];

    // 块解析完成后统一回填资源位置（表格单元格里的图也要能定位）
    assignAssetAnchors(blocks, ctx.assets);

    const metadata = { ...parseCoreProps(coreXml), ...parseAppProps(appXml) };
    const document: DocxDocument = {
        metadata: Object.keys(metadata).length ? metadata : undefined,
        section: body ? parseSection(body) : undefined,
        blocks
    };
    if (ctx.assets.length) document.assets = ctx.assets;

    return { document, messages, text: blocksToText(blocks) };
}
