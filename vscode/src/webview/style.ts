// Webview 内联基础样式（布局骨架）。幻灯片的具体背景/字体/主题色来自
// pptxToHtml 返回的 styles.global，由 main.ts 注入到 #pptx-styles。
export const baseCss = `
:root { --bg: #f6f8fa; --fg: #1f2328; --bd: #d0d7de; --pri: #0969da; --mut: #656d76; }
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: var(--bg); color: var(--fg); font: 14px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif; }
#app { padding: 16px; }
.toolbar { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }
.tbtn { padding: 4px 12px; border: 1px solid var(--bd); background: #fff; color: var(--fg); border-radius: 6px; cursor: pointer; font: inherit; font-size: 13px; }
.tbtn:hover { border-color: var(--pri); color: var(--pri); }
.tbtn.active { background: var(--pri); border-color: var(--pri); color: #fff; }
.hint { color: var(--mut); font-size: 12px; margin-top: 8px; }
.error { color: #cf222e; white-space: pre-wrap; }
#pptx-styles { display: none; }

/* PPTX 幻灯片 */
.pptx-slides { background: #fff; border: 1px solid var(--bd); border-radius: 8px; padding: 8px; }
.slide-scaler { display: inline-block; vertical-align: top; margin: 0 0 14px; }
/* 隔离本页全局样式对幻灯片渲染的影响（同 jdoc/examples 的做法） */
:where(.pptx-slides, .pptx-slides *) { box-sizing: content-box; line-height: normal; }

/* DOCX / XLSX 文档体 */
.docx-body, .xlsx-body { background: #fff; border: 1px solid var(--bd); border-radius: 8px; padding: 24px; max-width: 960px; }
.xlsx-body { padding: 8px; overflow: auto; }

/* PDF 页面 */
.pdf-page { background: #fff; border: 1px solid var(--bd); border-radius: 4px; margin: 0 auto 16px; box-shadow: 0 1px 3px rgba(0,0,0,.1); padding: 16px; }
.pdf-page .page-text { white-space: pre-wrap; word-break: break-word; }
`;
