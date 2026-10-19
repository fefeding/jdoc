// Webview 内联基础样式（布局骨架）。幻灯片的具体背景/字体/主题色来自
// pptxToHtml 返回的 styles.global，由 main.ts 注入到 #pptx-styles。
// 顶栏/分段切换/状态栏的统一样式见 chrome.ts（CHROME_CSS）。
export const baseCss = `
* { box-sizing: border-box; }
html, body, #app { height: 100%; margin: 0; padding: 0; }
body {
  background: var(--jc-bg, #f3f4f6);
  color: var(--jc-text, #1f2328);
  font: 14px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
}
#pptx-styles { display: none; }
.jc-app { display: flex; flex-direction: column; height: 100%; }
.jc-app > .jc-content { padding: 16px; }
.hint { color: var(--jc-muted, #656d76); font-size: 12px; padding: 24px; text-align: center; }
.error { color: #cf222e; white-space: pre-wrap; }

/* PPTX 幻灯片 */
.pptx-slides { display: flex; flex-direction: column; align-items: center; gap: 16px; }
.slide-scaler { display: inline-block; vertical-align: top; }
/* 隔离本页全局样式对幻灯片渲染的影响（同 jdoc/examples 的做法） */
::where(.pptx-slides, .pptx-slides *) { box-sizing: content-box; line-height: normal; }

/* DOCX / XLSX 文档体（查看器兜底路径，office-editor 有自己的样式） */
.docx-body, .xlsx-body { background: var(--jc-surface, #fff); border: 1px solid var(--jc-border, #d0d7de); border-radius: 10px; padding: 24px; max-width: 960px; margin: 0 auto; }
.xlsx-body { padding: 8px; overflow: auto; }

/* PDF 页面 */
.pdf-pages { display: flex; flex-direction: column; align-items: center; }
.pdf-page { background: #fff; border: 1px solid var(--jc-border, #d0d7de); border-radius: 4px; margin: 0 auto 16px; box-shadow: 0 1px 3px rgba(0,0,0,.1); padding: 16px; }
.pdf-page .page-text { white-space: pre-wrap; word-break: break-word; }
/* PDF 栅格化页面（canvas 渲染，兼容纯图片/扫描件） */
.pdf-canvas-page { padding: 0; overflow: hidden; position: relative; line-height: 0; }
.pdf-canvas-page canvas { display: block; width: 100%; height: 100%; }
.pdf-canvas-page.pending { display: flex; align-items: center; justify-content: center; color: var(--jc-muted, #656d76); font-size: 12px; line-height: 1.6; }
.pdf-canvas-page .pdf-text-layer { position: absolute; inset: 0; color: transparent; }
.pdf-canvas-page .pdf-text-layer span { position: absolute; white-space: pre; line-height: 1; font-family: sans-serif; transform-origin: 0 0; }

/* Markdown / 纯文本预览 */
:root {
  --jdoc-md-fg: var(--jc-text, #1f2328);
  --jdoc-md-bd: var(--jc-border, #d0d7de);
  --jdoc-md-mut: var(--jc-muted, #656d76);
  --jdoc-md-link: var(--jc-accent, #0969da);
  --jdoc-md-code-bg: #f6f8fa;
  --jdoc-md-th-bg: #f6f8fa;
  --jdoc-md-row-bg: rgba(246,248,250,.6);
}
.md-body, .txt-body { background: var(--jc-surface, #fff); border: 1px solid var(--jc-border, #d0d7de); border-radius: 10px; padding: 28px 32px; max-width: 960px; margin: 0 auto; }
.txt-body { padding: 16px 20px; overflow: auto; }
`;
