/**
 * Markdown 默认排版样式
 *
 * 作用域全部限定在 `.jdoc-md` 内，避免污染宿主页面；宿主可传 `includeStyle: false` 自行接管。
 *
 * @module markdown/to-html
 */
export const MARKDOWN_CSS = `
.jdoc-md { color: var(--jdoc-md-fg, #1f2328); font: 15px/1.75 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif; word-wrap: break-word; }
.jdoc-md > *:first-child { margin-top: 0; }
.jdoc-md > *:last-child { margin-bottom: 0; }
.jdoc-md h1, .jdoc-md h2, .jdoc-md h3, .jdoc-md h4, .jdoc-md h5, .jdoc-md h6 { margin: 24px 0 12px; font-weight: 600; line-height: 1.3; }
.jdoc-md h1 { font-size: 1.9em; padding-bottom: .3em; border-bottom: 1px solid var(--jdoc-md-bd, #d8dee4); }
.jdoc-md h2 { font-size: 1.5em; padding-bottom: .3em; border-bottom: 1px solid var(--jdoc-md-bd, #d8dee4); }
.jdoc-md h3 { font-size: 1.25em; }
.jdoc-md h4 { font-size: 1.05em; }
.jdoc-md h5, .jdoc-md h6 { font-size: 1em; color: var(--jdoc-md-mut, #656d76); }
.jdoc-md p, .jdoc-md blockquote, .jdoc-md ul, .jdoc-md ol, .jdoc-md dl, .jdoc-md table, .jdoc-md pre, .jdoc-md details { margin: 0 0 14px; }
.jdoc-md a { color: var(--jdoc-md-link, #0969da); text-decoration: none; }
.jdoc-md a:hover { text-decoration: underline; }
.jdoc-md ul, .jdoc-md ol { padding-left: 2em; }
.jdoc-md li + li { margin-top: .25em; }
.jdoc-md li > ul, .jdoc-md li > ol { margin: .25em 0 0; }
.jdoc-md blockquote { padding: 0 1em; color: var(--jdoc-md-mut, #656d76); border-left: .25em solid var(--jdoc-md-bd, #d8dee4); }
.jdoc-md code { padding: .2em .4em; margin: 0; font-size: 85%; background: var(--jdoc-md-code-bg, rgba(175,184,193,.2)); border-radius: 6px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }
.jdoc-md pre { padding: 14px 16px; overflow: auto; background: var(--jdoc-md-code-bg, #f6f8fa); border-radius: 8px; line-height: 1.5; }
.jdoc-md pre code { padding: 0; background: transparent; font-size: 90%; }
.jdoc-md table { border-collapse: collapse; display: block; width: max-content; max-width: 100%; overflow: auto; }
.jdoc-md th, .jdoc-md td { padding: 6px 13px; border: 1px solid var(--jdoc-md-bd, #d8dee4); }
.jdoc-md th { background: var(--jdoc-md-th-bg, #f6f8fa); font-weight: 600; }
.jdoc-md tr:nth-child(2n) td { background: var(--jdoc-md-row-bg, rgba(246,248,250,.6)); }
.jdoc-md img { max-width: 100%; }
.jdoc-md hr { height: .25em; margin: 24px 0; border: 0; background: var(--jdoc-md-bd, #d8dee4); }
.jdoc-md kbd { padding: .2em .4em; border: 1px solid var(--jdoc-md-bd, #d8dee4); border-radius: 6px; font-size: 85%; }
.jdoc-md del { color: var(--jdoc-md-mut, #656d76); }
.jdoc-md details { padding: 8px 12px; background: var(--jdoc-md-code-bg, #f6f8fa); border-radius: 8px; }
.jdoc-md summary { cursor: pointer; font-weight: 600; }
`;
