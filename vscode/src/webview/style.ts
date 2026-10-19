// Webview 内联基础样式（布局骨架）。幻灯片的具体背景/字体/主题色来自
// pptxToHtml 返回的 styles.global，由 main.ts 注入到 #pptx-styles。
// 顶栏/分段切换/状态栏的统一样式见 chrome.ts（CHROME_CSS）；
// PDF 阅读器（工具栏 + 缩略图栏）的样式见 pdf-viewer.ts（PDF_CSS）。
export const baseCss = `
* { box-sizing: border-box; }
html, body, #app { height: 100%; margin: 0; padding: 0; }
body {
  background: var(--jc-bg);
  color: var(--jc-text);
  font-family: var(--jc-font);
  font-size: var(--vscode-font-size, 13px);
  line-height: 1.6;
}
#pptx-styles { display: none; }
.jc-app { display: flex; flex-direction: column; height: 100%; }
.jc-app > .jc-content { padding: 16px; }
.hint { color: var(--jc-muted); font-size: 12px; padding: 24px; text-align: center; }
.error { color: var(--vscode-errorForeground, #f14c4c); white-space: pre-wrap; }

/* PPTX 幻灯片 */
.pptx-slides { display: flex; flex-direction: column; align-items: center; gap: 16px; }
.slide-scaler { display: inline-block; vertical-align: top; }
/* 隔离本页全局样式对幻灯片渲染的影响（同 jdoc/examples 的做法） */
::where(.pptx-slides, .pptx-slides *) { box-sizing: content-box; line-height: normal; }

/* DOCX / XLSX 文档体（查看器兜底路径，office-editor 有自己的样式） */
/* 文档纸面：DOCX/XLSX 的 HTML 自带深色文字与浅色表格，保持白底以保证可读性 */
.docx-body, .xlsx-body { background: #fff; color: #1f2328; border: 1px solid var(--jc-border-soft); border-radius: 10px; padding: 24px; max-width: 960px; margin: 0 auto; box-shadow: var(--jc-shadow, none); }
.xlsx-body { padding: 8px; overflow: auto; }

/* Markdown / 纯文本预览 */
/* 注意：这些令牌引用了 --jc-*（声明在 body 上），因此必须同样声明在 body 上——
   自定义属性里的 var() 在声明处解析，写在 :root 会因取不到 --jc-* 而整体失效。 */
body {
  --jdoc-md-fg: var(--jc-text);
  --jdoc-md-bd: var(--jc-border);
  --jdoc-md-mut: var(--jc-muted);
  --jdoc-md-link: var(--jc-link);
  --jdoc-md-code-bg: var(--jc-code-bg);
  --jdoc-md-th-bg: var(--jc-surface-2);
  --jdoc-md-row-bg: var(--jc-hover);
}
.md-body, .txt-body { background: var(--jc-surface); border-radius: 10px; padding: 28px 32px; max-width: 960px; margin: 0 auto; box-shadow: var(--jc-shadow, none); }
.txt-body { padding: 16px 20px; overflow: auto; }
.txt-body .jdoc-txt { font-family: var(--jc-font-mono); }
`;
