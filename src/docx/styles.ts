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
    childElements,
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
    type StyleRunProps
} from './style-props';

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
}

/** 构建空索引（styles.xml 缺失时的降级） */
export function emptyStylesIndex(): StylesIndex {
    return { byId: new Map(), byName: new Map(), defaultRun: {}, defaultParagraph: {} };
}

/**
 * 解析 word/styles.xml 为样式索引
 * @param xml - styles.xml 文本
 */
export function buildStylesIndex(xml: string): StylesIndex {
    const index = emptyStylesIndex();
    if (!xml) return index;

    const stylesRoot = findStylesRoot(parseXml(xml));
    if (!stylesRoot) return index;

    // 文档级默认属性 w:docDefaults
    const docDefaults = firstChild(stylesRoot, 'docDefaults');
    if (docDefaults) {
        const rPrDefault = firstChild(docDefaults, 'rPrDefault');
        if (rPrDefault) index.defaultRun = readRunProps(firstChild(rPrDefault, 'rPr'));
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
            run: readRunProps(firstChild(styleEl, 'rPr'))
        };
        rawById.set(id, style);
        if (!index.byName.has(style.name)) index.byName.set(style.name, id);
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
