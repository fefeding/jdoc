/**
 * DOCX 样式索引
 *
 * Word 样式通过 w:basedOn 层层继承，且「Heading 1」这类语义只存在于样式名（w:name）。
 * 本模块把 word/styles.xml 解析成按 styleId 查询、且已合并继承链的扁平结构。
 *
 * @module docx/styles
 */

import {
    attr,
    attrNumber,
    attrTrue,
    childElements,
    findFirst,
    firstChild,
    localName,
    parseXml,
    type XmlElement,
    type XmlNode
} from '../core/xml';
import {
    mergeParagraphProps,
    mergeRunProps,
    readParagraphProps,
    readRunProps,
    type StyleParagraphProps,
    type StyleRunProps,
    type ThemeFontSlot,
    type ThemeFonts
} from './style-props';

/**
 * 解析 theme1.xml 的主题字体。
 * rFonts 的 asciiTheme/eastAsiaTheme 引用 major/minor 槽位，需要这里换成真实字体名。
 */
export function parseThemeFonts(xml: string): ThemeFonts | undefined {    if (!xml) return undefined;
    const root = findFirst(parseXml(xml), 'theme');
    if (!root) return undefined;
    const themeElements = firstChild(root, 'themeElements');
    const scheme = themeElements ? firstChild(themeElements, 'fontScheme') : undefined;
    if (!scheme) return undefined;

    const readSlot = (container: XmlElement | undefined): ThemeFontSlot => {
        if (!container) return {};
        const latin = firstChild(container, 'latin');
        const ea = firstChild(container, 'ea');
        // a:ea 常为空串，中文实际字体挂在 <a:font script="Hans" typeface="宋体"/>
        let hans: string | undefined;
        for (const font of childElements(container, 'font')) {
            if ((attr(font, 'script') || '').toLowerCase() === 'hans') {
                hans = attr(font, 'typeface') || undefined;
                break;
            }
        }
        return {
            ascii: (latin && attr(latin, 'typeface')) || undefined,
            eastAsia: (ea && attr(ea, 'typeface')) || hans
        };
    };

    return {
        major: readSlot(firstChild(scheme, 'majorFont')),
        minor: readSlot(firstChild(scheme, 'minorFont'))
    };
}

/** 单个样式条目（已合并继承链） */
export interface DocxStyle {
    /** 样式 ID（正文 w:pStyle/@w:val 引用它） */
    id: string;
    /** 样式名，如 "heading 1"、"Title" */
    name: string;
    /** 基于哪个样式 */
    basedOn?: string;
    /** 最终生效的段落属性 */
    paragraph: StyleParagraphProps;
    /** 最终生效的字符属性 */
    run: StyleRunProps;
}

/** 样式索引 */
export interface StylesIndex {
    byId: Map<string, DocxStyle>;
    byName: Map<string, string>;
    /** 文档默认字符属性 */
    defaultRun: StyleRunProps;
    /** 文档默认段落属性 */
    defaultParagraph: StyleParagraphProps;
    /** 默认段落样式 ID（w:default="1" 的 paragraph 样式，未标 pStyle 的段落使用它） */
    defaultParagraphStyleId?: string;
}

/** 构建空索引（styles.xml 缺失时的降级） */
export function emptyStylesIndex(): StylesIndex {
    return { byId: new Map(), byName: new Map(), defaultRun: {}, defaultParagraph: {} };
}

/**
 * 解析 word/styles.xml 为样式索引
 * @param xml - styles.xml 文本
 * @param themeFonts - 主题字体（theme1.xml），用于还原 asciiTheme/eastAsiaTheme 引用
 */
export function buildStylesIndex(xml: string, themeFonts?: ThemeFonts): StylesIndex {
    const index = emptyStylesIndex();
    if (!xml) return index;

    const stylesRoot = findStylesRoot(parseXml(xml));
    if (!stylesRoot) return index;

    // 文档级默认属性 w:docDefaults
    const docDefaults = firstChild(stylesRoot, 'docDefaults');
    if (docDefaults) {
        const rPrDefault = firstChild(docDefaults, 'rPrDefault');
        if (rPrDefault) index.defaultRun = readRunProps(firstChild(rPrDefault, 'rPr'), themeFonts);
        const pPrDefault = firstChild(docDefaults, 'pPrDefault');
        if (pPrDefault) index.defaultParagraph = readParagraphProps(firstChild(pPrDefault, 'pPr'));
    }

    // 收集原始样式
    const rawById = new Map<string, DocxStyle>();
    for (const styleEl of childElements(stylesRoot, 'style')) {
        const id = attr(styleEl, 'styleId');
        if (!id) continue;
        const nameEl = firstChild(styleEl, 'name');
        const basedOnEl = firstChild(styleEl, 'basedOn');
        const style: DocxStyle = {
            id,
            name: (nameEl && attr(nameEl, 'val')) || id,
            basedOn: basedOnEl ? attr(basedOnEl, 'val') : undefined,
            paragraph: readParagraphProps(firstChild(styleEl, 'pPr')),
            run: readRunProps(firstChild(styleEl, 'rPr'), themeFonts)
        };
        rawById.set(id, style);
        if (!index.byName.has(style.name)) index.byName.set(style.name, id);
        // 未显式标 pStyle 的段落使用默认段落样式（Word 的 Normal）
        if (
            index.defaultParagraphStyleId === undefined &&
            (attr(styleEl, 'type') || 'paragraph') === 'paragraph' &&
            attrTrue(styleEl, 'default')
        ) {
            index.defaultParagraphStyleId = id;
        }
    }

    // 沿 basedOn 链自底向上合并
    const resolved = new Map<string, DocxStyle>();
    const resolve = (style: DocxStyle, seen: Set<string>): DocxStyle => {
        const cached = resolved.get(style.id);
        if (cached) return cached;
        if (seen.has(style.id)) return style; // 循环继承保护
        seen.add(style.id);

        let paragraph = { ...index.defaultParagraph };
        let run = { ...index.defaultRun };

        const parent = style.basedOn ? rawById.get(style.basedOn) : undefined;
        if (parent) {
            const merged = resolve(parent, seen);
            paragraph = mergeParagraphProps(paragraph, merged.paragraph);
            run = mergeRunProps(run, merged.run);
        }
        paragraph = mergeParagraphProps(paragraph, style.paragraph);
        run = mergeRunProps(run, style.run);

        const out: DocxStyle = { id: style.id, name: style.name, basedOn: style.basedOn, paragraph, run };
        resolved.set(style.id, out);
        return out;
    };

    for (const [id, style] of rawById) index.byId.set(id, resolve(style, new Set()));

    return index;
}

/** 定位 w:styles 根元素（容忍不同命名空间前缀） */
function findStylesRoot(roots: XmlNode[]): XmlElement | undefined {
    for (const node of roots) {
        if (node.type === 'element' && localName(node.tag) === 'styles') return node;
    }
    // 少数生成器省略外层 styles，直接给 w:style 列表
    for (const node of roots) {
        if (node.type === 'element' && childElements(node, 'style').length) return node;
    }
    return undefined;
}

/** 主题色板：schemeClr val -> #RRGGBB（dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink） */
export type ThemeColors = Map<string, string>;

/** schemeClr 别名：bg1/tx1/bg2/tx2 与 lt1/dk1/lt2/dk2 等价 */
const SCHEME_ALIASES: Record<string, string> = {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'lt2',
    tx2: 'dk2'
};

/** 解析 theme1.xml 的 clrScheme 主题色板 */
export function parseThemeColors(xml: string): ThemeColors {
    const colors: ThemeColors = new Map();
    if (!xml) return colors;
    const root = findFirst(parseXml(xml), 'theme');
    const themeElements = root ? firstChild(root, 'themeElements') : undefined;
    const scheme = themeElements ? firstChild(themeElements, 'clrScheme') : undefined;
    if (!scheme) return colors;

    for (const slot of childElements(scheme)) {
        const name = localName(slot.tag);
        const srgb = firstChild(slot, 'srgbClr');
        if (srgb) {
            const val = attr(srgb, 'val');
            if (val) colors.set(name, `#${val.replace(/^#/, '').toUpperCase()}`);
            continue;
        }
        const sys = firstChild(slot, 'sysClr');
        if (sys) {
            const last = attr(sys, 'lastClr');
            if (last) colors.set(name, `#${last.replace(/^#/, '').toUpperCase()}`);
        }
    }
    return colors;
}

/** RGB -> HSL（各分量 0-1） */
function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    let h: number;
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
    return [h, s, l];
}

function hueToRgb(p: number, q: number, t: number): number {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
}

/** HSL -> RGB（各分量 0-1） */
function hslToRgb(h: number, s: number, l: number): [number, number, number] {
    if (s === 0) return [l, l, l];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    return [hueToRgb(p, q, h + 1 / 3), hueToRgb(p, q, h), hueToRgb(p, q, h - 1 / 3)];
}

/** HSL 亮度修正（lumMod/lumOff 作用亮度，shade/tint 作用通道，单位 1/100000） */
function applyLumMods(hex: string, colorEl: XmlElement): string {
    let r = parseInt(hex.slice(1, 3), 16) / 255;
    let g = parseInt(hex.slice(3, 5), 16) / 255;
    let b = parseInt(hex.slice(5, 7), 16) / 255;

    const shade = firstChild(colorEl, 'shade');
    if (shade) {
        const k = attrNumber(shade, 'val', 100000) / 100000;
        r *= k;
        g *= k;
        b *= k;
    }
    const tint = firstChild(colorEl, 'tint');
    if (tint) {
        const k = attrNumber(tint, 'val', 0) / 100000;
        r = r * k + (1 - k);
        g = g * k + (1 - k);
        b = b * k + (1 - k);
    }

    const lumMod = firstChild(colorEl, 'lumMod');
    const lumOff = firstChild(colorEl, 'lumOff');
    if (lumMod || lumOff) {
        const mod = lumMod ? attrNumber(lumMod, 'val', 100000) / 100000 : 1;
        const off = lumOff ? attrNumber(lumOff, 'val', 0) / 100000 : 0;
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb(h, s, Math.min(1, Math.max(0, l * mod + off)));
    }

    const to = (v: number) => Math.round(Math.min(1, Math.max(0, v)) * 255)
        .toString(16)
        .padStart(2, '0');
    return `#${to(r)}${to(g)}${to(b)}`.toUpperCase();
}

/**
 * 解析 DrawingML 颜色元素（a:srgbClr / a:schemeClr，含 lumMod/lumOff/shade/tint 子元素）。
 * 用于文本框填充、边框等主题色引用的还原。
 */
export function resolveDrawingColor(
    colorEl: XmlElement | undefined,
    themeColors: ThemeColors
): string | undefined {
    if (!colorEl) return undefined;
    const name = localName(colorEl.tag);

    if (name === 'srgbClr') {
        const val = attr(colorEl, 'val');
        if (!val || val.toLowerCase() === 'auto') return undefined;
        const hex = `#${val.replace(/^#/, '').toUpperCase()}`;
        return /^[0-9A-F]{6}$/.test(hex.slice(1)) ? applyLumMods(hex, colorEl) : undefined;
    }

    if (name === 'schemeClr') {
        const val = attr(colorEl, 'val');
        if (!val) return undefined;
        const key = SCHEME_ALIASES[val] || val;
        const base = themeColors.get(key);
        if (!base) return undefined;
        return applyLumMods(base, colorEl);
    }

    return undefined;
}
