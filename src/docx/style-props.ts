/**
 * DOCX 属性读取（w:pPr / w:rPr）
 *
 * 单位统一换算：twips -> 磅（/20）、半磅 -> 磅（/2）。
 *
 * @module docx/style-props
 */

import { attr, attrNumber, attrTrue, firstChild, type XmlElement } from '../core/xml';

/** 主题字体（theme1.xml 的 majorFont/minorFont） */
export interface ThemeFontSlot {
    /** 西文字体（a:latin） */
    ascii?: string;
    /** 中文字体（a:ea，缺失时取 Hans script 字体） */
    eastAsia?: string;
}

export type ThemeFonts = { major: ThemeFontSlot; minor: ThemeFontSlot };

/** asciiTheme/eastAsiaTheme 值 -> 主题字体名（如 minorEastAsia -> 宋体） */
export function resolveThemeFont(
    themeFonts: ThemeFonts | undefined,
    themeAttr: string | undefined
): string | undefined {
    if (!themeFonts || !themeAttr) return undefined;
    const slot = themeAttr.toLowerCase().startsWith('major') ? themeFonts.major : themeFonts.minor;
    if (/eastasia|bidi|cs/i.test(themeAttr)) return slot.eastAsia || slot.ascii;
    return slot.ascii || slot.eastAsia;
}

/** 合并后的段落属性（w:pPr 子集） */
export interface StyleParagraphProps {
    align?: 'left' | 'center' | 'right' | 'justify';
    outlineLevel?: number;
    indentLeft?: number;
    indentRight?: number;
    /** 首行缩进；负值表示悬挂缩进 */
    indentFirstLine?: number;
    spacingBefore?: number;
    spacingAfter?: number;
    /** 行距倍数，1 为单倍行距 */
    lineSpacing?: number;
    isListParagraph?: boolean;
}

/** 合并后的字符属性（w:rPr 子集） */
export interface StyleRunProps {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    subscript?: boolean;
    superscript?: boolean;
    color?: string;
    /** 字号（磅） */
    size?: number;
    font?: string;
    highlight?: string;
}

/** w:jc 取值 -> 统一对齐名 */
export function normalizeAlign(value: string | undefined): StyleParagraphProps['align'] {
    if (!value) return undefined;
    const v = value.toLowerCase();
    if (v === 'center') return 'center';
    if (v === 'right' || v === 'end') return 'right';
    if (v === 'both' || v === 'distribute') return 'justify';
    if (v === 'left' || v === 'start') return 'left';
    return undefined;
}

/** Word 高亮枚举值 -> #RRGGBB */
function highlightToHex(value: string): string {
    const map: Record<string, string> = {
        yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF',
        blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080',
        darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
        darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0',
        black: '000000', white: 'FFFFFF'
    };
    return map[value] || 'FFFF00';
}

/** 从 w:rPr 读取字符属性 */
export function readRunProps(rPr: XmlElement | undefined, themeFonts?: ThemeFonts): StyleRunProps {
    if (!rPr) return {};
    const props: StyleRunProps = {};

    const b = firstChild(rPr, 'b');
    if (b) props.bold = attrTrue(b, 'val');
    const i = firstChild(rPr, 'i');
    if (i) props.italic = attrTrue(i, 'val');

    const u = firstChild(rPr, 'u');
    if (u) {
        // w:val="none" 为显式关闭下划线
        const style = attr(u, 'val');
        props.underline = !(style && style.toLowerCase() === 'none');
    }

    const strike = firstChild(rPr, 'strike') || firstChild(rPr, 'dstrike');
    if (strike) props.strike = attrTrue(strike, 'val');

    const color = firstChild(rPr, 'color');
    if (color) {
        const value = attr(color, 'val');
        // "auto" 表示自动色，不映射为具体颜色
        if (value && value.toLowerCase() !== 'auto') props.color = `#${value.replace(/^#/, '')}`;
    }

    const sz = firstChild(rPr, 'sz');
    if (sz) {
        const half = attrNumber(sz, 'val', 0);
        if (half > 0) props.size = half / 2;
    }

    const fonts = firstChild(rPr, 'rFonts');
    if (fonts) {
        // 字面属性优先；缺失时还原 asciiTheme/eastAsiaTheme 主题引用
        const eastAsia =
            attr(fonts, 'eastAsia') || resolveThemeFont(themeFonts, attr(fonts, 'eastAsiaTheme'));
        const ascii =
            attr(fonts, 'ascii') ||
            attr(fonts, 'hAnsi') ||
            resolveThemeFont(themeFonts, attr(fonts, 'asciiTheme')) ||
            resolveThemeFont(themeFonts, attr(fonts, 'hAnsiTheme'));
        // 中文取 eastAsia，西文取 ascii/hAnsi
        const name = eastAsia || ascii || attr(fonts, 'cs');
        if (name) props.font = name;
    }

    const highlight = firstChild(rPr, 'highlight');
    if (highlight) {
        const value = attr(highlight, 'val');
        if (value && value.toLowerCase() !== 'none') props.highlight = `#${highlightToHex(value)}`;
    }

    const vertAlign = firstChild(rPr, 'vertAlign');
    if (vertAlign) {
        const value = (attr(vertAlign, 'val') || '').toLowerCase();
        if (value === 'subscript') props.subscript = true;
        else if (value === 'superscript') props.superscript = true;
    }

    return props;
}

/** 从 w:pPr 读取段落属性 */
export function readParagraphProps(pPr: XmlElement | undefined): StyleParagraphProps {
    if (!pPr) return {};
    const props: StyleParagraphProps = {};

    const jc = firstChild(pPr, 'jc');
    if (jc) props.align = normalizeAlign(attr(jc, 'val'));

    const outline = firstChild(pPr, 'outlineLvl');
    if (outline) props.outlineLevel = attrNumber(outline, 'val', 9);

    const ind = firstChild(pPr, 'ind');
    if (ind) {
        const left = attrNumber(ind, 'left', NaN);
        const start = attrNumber(ind, 'start', NaN);
        const leftTwips = Number.isFinite(left) ? left : start;
        if (Number.isFinite(leftTwips)) props.indentLeft = leftTwips / 20;

        const right = attrNumber(ind, 'right', NaN);
        const end = attrNumber(ind, 'end', NaN);
        const rightTwips = Number.isFinite(right) ? right : end;
        if (Number.isFinite(rightTwips)) props.indentRight = rightTwips / 20;

        // firstLine 与 hanging 互斥，同时出现时 hanging（悬挂缩进）优先
        const firstLine = attrNumber(ind, 'firstLine', NaN);
        const hanging = attrNumber(ind, 'hanging', NaN);
        if (Number.isFinite(hanging)) props.indentFirstLine = -hanging / 20;
        else if (Number.isFinite(firstLine)) props.indentFirstLine = firstLine / 20;
    }

    const spacing = firstChild(pPr, 'spacing');
    if (spacing) {
        const before = attrNumber(spacing, 'before', NaN);
        const after = attrNumber(spacing, 'after', NaN);
        if (Number.isFinite(before)) props.spacingBefore = before / 20;
        if (Number.isFinite(after)) props.spacingAfter = after / 20;

        const line = attrNumber(spacing, 'line', NaN);
        if (Number.isFinite(line) && line > 0) {
            const rule = (attr(spacing, 'lineRule') || 'auto').toLowerCase();
            // auto/atLeast 以 240 = 单倍行距为基数；exact 已是磅值
            props.lineSpacing = rule === 'exact' ? line / 20 : line / 240;
        }
    }

    if (firstChild(pPr, 'numPr')) props.isListParagraph = true;

    return props;
}

/** 合并字符属性：override 中非 undefined 的值覆盖 base */
export function mergeRunProps(base: StyleRunProps, override: StyleRunProps): StyleRunProps {
    const out: StyleRunProps = { ...base };
    for (const key of Object.keys(override)) {
        const value = (override as Record<string, unknown>)[key];
        if (value !== undefined) (out as Record<string, unknown>)[key] = value;
    }
    return out;
}

/** 合并段落属性 */
export function mergeParagraphProps(
    base: StyleParagraphProps,
    override: StyleParagraphProps
): StyleParagraphProps {
    const out: StyleParagraphProps = { ...base };
    for (const key of Object.keys(override)) {
        const value = (override as Record<string, unknown>)[key];
        if (value !== undefined) (out as Record<string, unknown>)[key] = value;
    }
    return out;
}

/** 样式名 -> 标题层级；undefined 表示非标题样式 */
export function headingLevelOf(styleName: string | undefined): 1 | 2 | 3 | 4 | 5 | 6 | undefined {
    if (!styleName) return undefined;
    const name = styleName.toLowerCase().trim();

    const match = /^heading\s*([1-9])$/.exec(name) || /^标题\s*([1-9])$/.exec(name);
    if (match) {
        const level = Number(match[1]);
        return level <= 6 ? (level as 1 | 2 | 3 | 4 | 5 | 6) : 6;
    }
    if (name === 'title' || name === '标题') return 1;
    if (name === 'subtitle' || name === '副标题') return 2;
    return undefined;
}

/** 样式名 -> 引用块；undefined 表示不是引用样式 */
export function quoteStyleOf(styleName: string | undefined): boolean {
    if (!styleName) return false;
    const name = styleName.toLowerCase().trim();
    return name === 'quote' || name === 'intense quote' || name === 'blockquote' || name === '引用';
}

/**
 * 样式名 -> 目录条目层级；undefined 表示不是目录条目样式
 *
 * Word/WPS 的目录条目使用 `toc 1` / `toc 2` / `toc 3` 段落样式（中文版可能叫
 * `目录 1`）。据此可把一串缓存段落折叠成单个 toc 块，还原带点线导引的目录。
 */
export function tocLevelOf(styleName: string | undefined): number | undefined {
    if (!styleName) return undefined;
    const match = /^toc\s*([1-9])$/.exec(styleName.toLowerCase().trim()) || /^目录\s*([1-9])$/.exec(styleName.trim());
    if (!match) return undefined;
    return Number(match[1]);
}
