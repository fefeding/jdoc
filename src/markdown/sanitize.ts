/**
 * Markdown 渲染结果的 HTML 清洗
 *
 * Markdown 允许内嵌原始 HTML，直接 innerHTML 会带进 script / iframe / on* 事件等风险。
 * 这里做一层白名单清洗：只保留排版需要的标签与属性，并校验链接协议。
 * Webview 场景有 CSP 兜底，但库也要能安全用于任意宿主（后台预览、邮件正文等）。
 *
 * @module markdown/sanitize
 */

/** 允许的标签（其余标签会被丢弃标签本身、保留内部文本） */
const ALLOWED_TAGS = new Set([
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'p', 'br', 'hr', 'blockquote', 'pre', 'code', 'span', 'div',
    'strong', 'b', 'em', 'i', 'del', 's', 'u', 'mark', 'sub', 'sup', 'small', 'kbd',
    'ul', 'ol', 'li', 'dl', 'dt', 'dd',
    'table', 'thead', 'tbody', 'tfoot', 'tr', 'th', 'td', 'caption', 'colgroup', 'col',
    'a', 'img', 'figure', 'figcaption',
    'details', 'summary', 'section', 'article', 'header', 'footer', 'nav', 'aside', 'main'
]);

/** 连内容一并丢弃的标签（脚本、样式、嵌入、表单控件） */
const DROP_CONTENT_TAGS = new Set([
    'script', 'style', 'iframe', 'object', 'embed', 'applet', 'link', 'meta', 'base',
    'form', 'input', 'button', 'textarea', 'select', 'option', 'noscript', 'template', 'svg', 'math'
]);

/** 各标签允许保留的属性 */
const ALLOWED_ATTRS: Record<string, Set<string>> = {
    a: new Set(['href', 'title', 'target', 'rel']),
    img: new Set(['src', 'alt', 'title', 'width', 'height', 'loading']),
    th: new Set(['colspan', 'rowspan', 'align', 'style']),
    td: new Set(['colspan', 'rowspan', 'align', 'style']),
    col: new Set(['span', 'style']),
    colgroup: new Set(['span']),
    ol: new Set(['start']),
    code: new Set(['class']),
    pre: new Set(['class']),
    div: new Set(['class', 'align']),
    span: new Set(['class', 'style']),
    details: new Set(['open']),
    // data-tex / data-diagram 是 Markdown 增强渲染（KaTeX / Mermaid）的原始源码载体
    '*': new Set([
        'class', 'id', 'align', 'style', 'aria-hidden', 'colspan', 'rowspan', 'start', 'open', 'type',
        'data-tex', 'data-diagram'
    ])
};

/** 安全的 URL 协议 / 前缀 */
const SAFE_URL = /^(?:https?:|mailto:|tel:|ftp:|#|\/|\.\/|\.\.\/|data:image\/)/i;
/** 相对路径（不含协议前缀）也视为安全，如 `./a.png`、`api/doc` */
const RELATIVE_URL = /^[^:?#]*[?#]?[^:]*$/;

/** 判断 URL 是否安全（拒绝 javascript: / vbscript: / data:text-html 等） */
function isSafeUrl(value: string): boolean {
    const url = value.trim();
    if (!url) return true;
    if (/^[a-z][a-z0-9+.-]*:/i.test(url)) return SAFE_URL.test(url);
    return RELATIVE_URL.test(url) || SAFE_URL.test(url);
}

/** HTML 文本/属性转义（allowHtml 关闭时把原始 HTML 当纯文本显示） */
export function escapeHtml(value: string): string {
    return value.replace(/[&<>"']/g, (c) => {
        switch (c) {
            case '&': return '&amp;';
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '"': return '&quot;';
            default: return '&#39;';
        }
    });
}

/** 属性值转义（统一用双引号包裹输出） */
function escapeAttr(value: string): string {
    return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * 清洗 HTML 字符串
 * @param html - 待清洗的 HTML（通常来自 Markdown 渲染结果）
 * @returns 清洗后的 HTML
 */
export function sanitizeHtml(html: string): string {
    if (!html) return '';
    // 去注释与文档类型声明
    let source = html.replace(/<!--[\s\S]*?-->/g, '').replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '').replace(/<!(?:DOCTYPE|doctype)[^>]*>/g, '');

    let out = '';
    // 处在"丢弃内容"区间时跳过的标签名
    let skipTag: string | null = null;
    const tagRe = /<(\/)?([a-zA-Z][a-zA-Z0-9-]*)((?:\s+[^>]*?)?)(\/)?>/g;
    let last = 0;
    let m: RegExpExecArray | null;

    while ((m = tagRe.exec(source))) {
        const [full, closing, rawTag, rawAttrs, selfClose] = m;
        const tag = rawTag.toLowerCase();

        out += source.slice(last, m.index);
        last = m.index + full.length;

        if (skipTag) {
            if (closing && tag === skipTag) skipTag = null;
            continue;
        }
        if (!closing && DROP_CONTENT_TAGS.has(tag)) {
            skipTag = tag;
            continue;
        }
        if (closing) {
            if (ALLOWED_TAGS.has(tag)) out += `</${tag}>`;
            continue;
        }
        if (!ALLOWED_TAGS.has(tag)) continue;

        // 属性白名单
        const allowed = ALLOWED_ATTRS[tag] || ALLOWED_ATTRS['*'];
        const common = ALLOWED_ATTRS['*'];
        const attrs: string[] = [];
        const attrRe = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'`=<>]+))/g;
        let a: RegExpExecArray | null;
        while ((a = attrRe.exec(rawAttrs || ''))) {
            const name = a[1].toLowerCase();
            const value = a[2] ?? a[3] ?? a[4] ?? '';
            if (!allowed.has(name) && !common.has(name)) continue;
            if ((name === 'href' || name === 'src') && !isSafeUrl(value)) continue;
            if (name === 'style' && /expression|url\s*\(/i.test(value)) continue;
            attrs.push(` ${name}="${escapeAttr(value)}"`);
        }
        if (tag === 'a' && !/ rel=/i.test(rawAttrs || '')) attrs.push(' rel="noopener noreferrer"');
        out += `<${tag}${attrs.join('')}${selfClose ? ' /' : ''}>`;
    }
    out += source.slice(last);

    return out;
}

/**
 * HTML -> 纯文本（去标签 + 反转义常用实体），用于字数统计与纯文本提取
 */
export function htmlToText(html: string): string {
    return html
        .replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(p|div|li|tr|h[1-6]|blockquote|pre)>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&amp;/g, '&')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}
