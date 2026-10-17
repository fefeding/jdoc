/**
 * 零依赖 XML 解析器
 *
 * DOCX 需要直接读取 OOXML 部件，本模块提供最小可用的 XML → 节点树转换，
 * 并附带高频查询辅助函数，避免业务侧到处手写 for 循环。
 *
 * @module core/xml
 */

/** XML 元素节点 */
export interface XmlElement {
    type: 'element';
    /** 标签名（保留命名空间前缀，如 w:p） */
    tag: string;
    /** 属性表（取值已解码实体） */
    attrs: Record<string, string>;
    children: XmlNode[];
}

/** XML 文本节点 */
export interface XmlText {
    type: 'text';
    value: string;
}

/** XML 节点联合类型 */
export type XmlNode = XmlElement | XmlText;

/** 判断节点是否为元素 */
export function isElement(node: XmlNode | undefined | null): node is XmlElement {
    return !!node && node.type === 'element';
}

const XML_ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** 解码 XML 实体（含数字字符引用） */
export function decodeXmlEntities(value: string): string {
    if (!value || value.indexOf('&') === -1) return value;
    return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (match, code: string) => {
        if (code.charAt(0) === '#') {
            const isHex = code.charAt(1) === 'x' || code.charAt(1) === 'X';
            const num = parseInt(isHex ? code.slice(2) : code.slice(1), isHex ? 16 : 10);
            if (Number.isNaN(num) || num < 0 || num > 0x10ffff) return match;
            try {
                return String.fromCodePoint(num);
            } catch {
                return match;
            }
        }
        const named = XML_ENTITIES[code.toLowerCase()];
        return named === undefined ? match : named;
    });
}

/** 去除命名空间前缀：w:p -> p */
export function localName(tag: string): string {
    const idx = tag.indexOf(':');
    return idx === -1 ? tag : tag.slice(idx + 1);
}

/** 找到标签中未被引号包裹的 '>' */
function findTagEnd(xml: string, from: number): number {
    let quote = '';
    for (let i = from; i < xml.length; i++) {
        const ch = xml[i];
        if (quote) {
            if (ch === quote) quote = '';
        } else if (ch === '"' || ch === "'") quote = ch;
        else if (ch === '>') return i;
    }
    return -1;
}

/** 找到标签名结束位置（空白或 / 之前） */
function findNameEnd(body: string): number {
    for (let i = 0; i < body.length; i++) {
        const ch = body[i];
        if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === '/') return i;
    }
    return body.length;
}

/** 解析标签内属性串 */
function parseXmlAttrs(raw: string): Record<string, string> {
    const attrs: Record<string, string> = {};
    if (!raw.trim()) return attrs;
    const re = /([^\s=/]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(raw)) !== null) {
        attrs[m[1]] = decodeXmlEntities(m[2] ?? m[3] ?? m[4] ?? '');
    }
    return attrs;
}

/**
 * 解析 XML 字符串
 * @param xml - XML 文本
 * @returns 顶层节点数组（XML 声明 / 注释 / DOCTYPE 会被跳过）
 */
export function parseXml(xml: string): XmlNode[] {
    const roots: XmlNode[] = [];
    if (!xml) return roots;

    const stack: XmlElement[] = [];
    let i = 0;
    const len = xml.length;

    const appendText = (value: string) => {
        if (!value) return;
        const top = stack[stack.length - 1];
        // 顶层纯空白（XML 声明与根元素之间的换行缩进）直接丢弃，
        // 否则调用方拿到的 roots 第一个节点往往是空白文本而非根元素
        if (!top && !value.trim()) return;
        const node: XmlText = { type: 'text', value: decodeXmlEntities(value) };
        if (top) top.children.push(node);
        else roots.push(node);
    };

    while (i < len) {
        const lt = xml.indexOf('<', i);
        if (lt === -1) {
            appendText(xml.slice(i));
            break;
        }
        if (lt > i) appendText(xml.slice(i, lt));

        // CDATA 原样保留，不解码实体
        if (xml.startsWith('<![CDATA[', lt)) {
            const end = xml.indexOf(']]>', lt);
            if (end === -1) {
                appendText(xml.slice(lt + 9));
                break;
            }
            const top = stack[stack.length - 1];
            const node: XmlText = { type: 'text', value: xml.slice(lt + 9, end) };
            if (top) top.children.push(node);
            else roots.push(node);
            i = end + 3;
            continue;
        }

        // 注释 / 处理指令 / DOCTYPE：整段跳过
        if (xml.startsWith('<!--', lt)) {
            const end = xml.indexOf('-->', lt);
            i = end === -1 ? len : end + 3;
            continue;
        }
        if (xml.startsWith('<?', lt)) {
            const end = xml.indexOf('?>', lt);
            i = end === -1 ? len : end + 2;
            continue;
        }
        if (xml.startsWith('<!', lt)) {
            let end = xml.indexOf('>', lt);
            const bracket = xml.indexOf('[', lt);
            if (bracket !== -1 && bracket < end) {
                const bracketEnd = xml.indexOf(']', bracket);
                if (bracketEnd !== -1) end = xml.indexOf('>', bracketEnd);
            }
            i = end === -1 ? len : end + 1;
            continue;
        }

        // 闭合标签
        if (xml.charAt(lt + 1) === '/') {
            const end = xml.indexOf('>', lt);
            if (end === -1) break;
            const tag = xml.slice(lt + 2, end).trim();
            for (let s = stack.length - 1; s >= 0; s--) {
                if (stack[s].tag === tag) {
                    stack.length = s;
                    break;
                }
            }
            i = end + 1;
            continue;
        }

        // 开标签
        const gt = findTagEnd(xml, lt + 1);
        if (gt === -1) break;

        const raw = xml.slice(lt + 1, gt);
        const selfClosing = raw.endsWith('/');
        const body = selfClosing ? raw.slice(0, -1) : raw;
        const nameEnd = findNameEnd(body);
        const element: XmlElement = {
            type: 'element',
            tag: body.slice(0, nameEnd).trim(),
            attrs: parseXmlAttrs(body.slice(nameEnd)),
            children: []
        };

        const top = stack[stack.length - 1];
        if (top) top.children.push(element);
        else roots.push(element);

        if (!selfClosing) stack.push(element);
        i = gt + 1;
    }

    return roots;
}

/**
 * 取子元素。
 * - 传入单个元素：返回它的直接子元素中匹配 tag 的那些
 * - 传入数组：返回数组中匹配 tag 的那些
 * @param node - 容器元素或节点数组
 * @param tag - 标签名，可省略命名空间前缀；省略则不过滤
 */
export function childElements(node: XmlNode | XmlNode[] | undefined | null, tag?: string): XmlElement[] {
    if (!node) return [];

    // 数组按「节点池」处理；单个元素按「容器」处理，取其 children
    const pool: XmlNode[] = Array.isArray(node) ? node : isElement(node) ? node.children : [node];

    const out: XmlElement[] = [];
    for (const item of pool) {
        if (!isElement(item)) continue;
        if (tag !== undefined && item.tag !== tag && localName(item.tag) !== tag) continue;
        out.push(item);
    }
    return out;
}

/** 取第一个匹配的直接子元素 */
export function firstChild(node: XmlNode | XmlNode[] | undefined | null, tag: string): XmlElement | undefined {
    return childElements(node, tag)[0];
}

/** 递归查找所有后代元素 */
export function findAll(node: XmlNode | XmlNode[] | undefined | null, tag: string): XmlElement[] {
    const out: XmlElement[] = [];
    const walk = (n: XmlNode | XmlNode[] | undefined | null) => {
        for (const item of Array.isArray(n) ? n : n ? [n] : []) {
            if (!isElement(item)) continue;
            if (item.tag === tag || localName(item.tag) === tag) out.push(item);
            walk(item.children);
        }
    };
    walk(node);
    return out;
}

/** 递归查找第一个后代元素 */
export function findFirst(node: XmlNode | XmlNode[] | undefined | null, tag: string): XmlElement | undefined {
    for (const item of Array.isArray(node) ? node : node ? [node] : []) {
        if (!isElement(item)) continue;
        if (item.tag === tag || localName(item.tag) === tag) return item;
        const nested = findFirst(item.children, tag);
        if (nested) return nested;
    }
    return undefined;
}

/** 读取属性（兼容无前缀写法） */
export function attr(node: XmlElement | undefined | null, name: string): string | undefined {
    if (!node) return undefined;
    const direct = node.attrs[name];
    if (direct !== undefined) return direct;
    for (const key of Object.keys(node.attrs)) {
        if (localName(key) === name) return node.attrs[key];
    }
    return undefined;
}

/** 读取属性并转数字，缺失或非法时返回默认值 */
export function attrNumber(node: XmlElement | undefined | null, name: string, fallback = 0): number {
    const raw = attr(node, name);
    if (raw === undefined) return fallback;
    const num = Number(raw);
    return Number.isFinite(num) ? num : fallback;
}

/**
 * 判断布尔属性（OOXML 中 "1"/"true"/"on" 均为真，缺省为真）
 */
export function attrTrue(node: XmlElement | undefined | null, name: string, defaultValue = true): boolean {
    if (!node) return defaultValue;
    const raw = attr(node, name);
    if (raw === undefined) return defaultValue;
    const v = String(raw).toLowerCase();
    return !(v === '0' || v === 'false' || v === 'off');
}

/** 提取节点内全部文本 */
export function textOf(node: XmlNode | XmlNode[] | undefined | null): string {
    if (!node) return '';
    let out = '';
    for (const item of Array.isArray(node) ? node : [node]) {
        if (item.type === 'text') out += item.value;
        else out += textOf(item.children);
    }
    return out;
}