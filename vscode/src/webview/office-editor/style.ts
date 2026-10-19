// office-editor 样式：外壳（工具栏/内容/状态栏）+ 块式 DOCX + 网格 XLSX。
// 顶栏/分段切换的统一样式见 chrome.ts（CHROME_CSS），这里只保留编辑器专属规则。

export const officeCss = `
/* 与共享外观层保持同一套令牌，便于深浅色一致。
   注意：引用了 body 上的 --jc-*，因此必须声明在 body 上（var() 在声明处解析）。 */
body {
  --oe-bg: var(--jc-bg, #f3f4f6);
  --oe-surface: var(--jc-surface, #ffffff);
  --oe-border: var(--jc-border, #e2e5ea);
  --oe-text: var(--jc-text, #1f2430);
  --oe-muted: var(--jc-muted, #6b7280);
  --oe-accent: var(--jc-accent, #2b6cb0);
  --oe-accent-soft: var(--jc-accent-soft, #e8f0fb);
  --oe-danger: #d64545;
  --oe-toolbar: var(--jc-surface-2, #fafbfc);
  --oe-shadow: 0 1px 2px rgba(16,24,40,.06), 0 1px 3px rgba(16,24,40,.1);
}
* { box-sizing: border-box; }
html, body, #app { height: 100%; margin: 0; }
body {
  font-family: var(--jc-font);
  font-size: var(--vscode-font-size, 13px);
  line-height: 1.5;
  color: var(--oe-text);
  background: var(--oe-bg);
}
.oe { display: flex; flex-direction: column; height: 100%; }
.oe-body { flex: 1; display: flex; min-height: 0; }
.oe-toolbar { display: flex; flex-wrap: wrap; align-items: center; gap: 4px; padding: 6px 12px; background: var(--oe-toolbar); border-bottom: 1px solid var(--oe-border); min-height: 42px; }
.oe-toolbar:empty { display: none; }
.oe-tb { appearance: none; border: 1px solid transparent; background: transparent; color: var(--oe-text); height: 30px; min-width: 30px; padding: 0 7px; border-radius: 6px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 5px; font-size: 13px; }
.oe-tb:hover { background: var(--oe-accent-soft); }
.oe-tb.active { background: var(--oe-accent); color: #fff; border-color: var(--oe-accent); }
.oe-tb:disabled { opacity: .4; cursor: default; background: transparent; }
.oe-tb svg { width: 17px; height: 17px; }
.oe-tb.oe-glyph { font-weight: 700; }
.oe-sep { width: 1px; align-self: stretch; margin: 4px 4px; background: var(--oe-border); }
/* 弹性占位：把工具栏后续按钮推到最右（docx.ts / xlsx.ts 的工具栏末尾在用） */
.oe-spacer { flex: 1 1 auto; min-width: 8px; }
.oe-select { height: 30px; border: 1px solid var(--oe-border); border-radius: 6px; background: var(--oe-surface); color: var(--oe-text); padding: 0 6px; font-size: 13px; cursor: pointer; }
.oe-select option { background: var(--vscode-dropdown-background, var(--oe-surface)); color: var(--vscode-dropdown-foreground, var(--oe-text)); }
.oe-toolbar input[type="text"], .oe-toolbar input[type="number"], .oe-toolbar textarea { background: var(--vscode-input-background, var(--oe-surface)); color: var(--vscode-input-foreground, var(--oe-text)); border: 1px solid var(--vscode-input-border, var(--oe-border)); border-radius: 6px; padding: 4px 8px; font-size: 13px; }
.oe-color { width: 30px; height: 30px; padding: 0; border: 1px solid var(--oe-border); border-radius: 6px; background: var(--oe-surface); cursor: pointer; }

.oe-content { flex: 1; overflow: auto; position: relative; min-width: 0; }
.oe-statusbar { display: flex; align-items: center; gap: 12px; padding: 5px 14px; background: var(--jc-statusbar); border-top: 1px solid var(--jc-border); font-size: 12px; color: var(--jc-statusbar-fg); }
.oe-statusbar .grow { flex: 1; }

/* ---------- 左侧目录 / 编号索引栏 ---------- */
.oe-outline { width: 232px; flex: 0 0 232px; display: flex; flex-direction: column; min-height: 0; background: var(--oe-surface); border-right: 1px solid var(--oe-border); }
.oe-outline.empty { display: none; }
.oe-outline.collapsed { width: 34px; flex-basis: 34px; }
.oe-outline.collapsed .oe-outline-title, .oe-outline.collapsed .oe-outline-body { display: none; }
.oe-outline.collapsed .oe-outline-head { justify-content: center; padding: 0; }
.oe-outline-head { display: flex; align-items: center; gap: 6px; padding: 8px 10px; border-bottom: 1px solid var(--oe-border); }
.oe-outline-title { font-size: 12px; font-weight: 600; letter-spacing: .04em; color: var(--oe-muted); }
.oe-outline-toggle { appearance: none; border: 0; background: transparent; cursor: pointer; padding: 2px 4px; border-radius: 5px; color: var(--oe-muted); line-height: 1; }
.oe-outline-toggle:hover { background: var(--oe-accent-soft); color: var(--oe-text); }
.oe-outline-toggle::before { content: '☰'; font-size: 13px; }
.oe-outline-body { flex: 1; min-height: 0; overflow: auto; padding: 6px 6px 24px; }
.oe-outline-list { display: flex; flex-direction: column; gap: 1px; }
.oe-outline-item { appearance: none; border: 0; background: transparent; text-align: left; cursor: pointer; border-radius: 6px; padding: 4px 8px; color: var(--oe-text); font-size: 12.5px; line-height: 1.45; display: flex; align-items: baseline; gap: 2px; }
.oe-outline-item:hover { background: var(--oe-accent-soft); }
/* 选中态：背景/前景用 VS Code 为列表提供的配对令牌。
   原来这里用 accent 文字压 accent-soft 底，两者对比度无保证（实测低至 1.33:1）。 */
.oe-outline-item.active { background: var(--jc-active); color: var(--jc-active-fg); font-weight: 600; box-shadow: inset 2px 0 0 var(--oe-accent); }
.oe-outline-item[data-level="1"] { font-weight: 600; }
.oe-outline-item[data-level="3"] { font-size: 12px; color: var(--oe-muted); }
.oe-outline-item[data-level="4"], .oe-outline-item[data-level="5"], .oe-outline-item[data-level="6"] { font-size: 11.5px; color: var(--oe-muted); }
.oe-outline-item[data-level="2"] { padding-left: 20px; }
.oe-outline-item[data-level="3"] { padding-left: 32px; }
.oe-outline-item[data-level="4"] { padding-left: 44px; }
.oe-outline-item[data-level="5"], .oe-outline-item[data-level="6"] { padding-left: 56px; }
.oe-outline-text { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.oe-outline-num { flex: 0 0 auto; font-variant-numeric: tabular-nums; opacity: .75; margin-right: .6em; }

/* ---------- 分页纸面 ---------- */
.oe-view-paged, .oe-doc { padding: 24px 16px 64px; }
.oe-doc { display: flex; flex-direction: column; align-items: center; }
/* 左侧留出块操作手柄的位置：.oe-block .oe-grip 是 left:-30px 的浮层 */
.oe-pages { display: flex; flex-direction: column; align-items: center; gap: 20px; padding-left: 34px; }
/* 纸面恒为白底深字，不随主题染色。
   原因：文档是「印在纸上」的内容，jdoc 自带 .docx-body{color:#1f2328}（近黑），
   它是后代上的类选择器，优先级高于 .oe-page 的 color 继承。若深色主题把纸面染暗，
   就会变成暗底 + 暗色正文，几乎不可见。与 Word / 浏览器 PDF 阅读器一致：
   纸白字黑，其余 UI（工具栏、目录栏、状态栏）才跟随主题。 */
.oe-page { position: relative; background: #fff; color: #1f2328; box-shadow: var(--oe-shadow); border-radius: 3px; box-sizing: border-box; overflow: visible; }
.oe-page-body { box-sizing: border-box; color: #1f2328; }
/* 单块高于整页（长表格 / 大代码框）时允许溢出，块内可滚动 */
.oe-tall { overflow: visible; }
.oe-page-num { position: absolute; left: 0; right: 0; bottom: 14px; text-align: center; font-size: 10px; color: #9aa1ac; user-select: none; pointer-events: none; }
/* 大纲跳转时短暂高亮目标标题 */
.oe-flash { animation: oe-flash .9s ease-out; }
@keyframes oe-flash { from { background: rgba(43,108,176,.18); } to { background: transparent; } }

/* ---------- 查看模式（DOCX/XLSX 只读预览） ---------- */
.oe-view { padding: 24px; display: flex; justify-content: center; align-items: flex-start; }
.oe-view .page { background: #fff; color: #1f2328; box-shadow: var(--oe-shadow); border-radius: 4px; padding: 48px 56px; max-width: 820px; width: 100%; }
.oe-view .page table { border-collapse: collapse; }
.oe-view img { max-width: 100%; }
/* 分页后每页自带纸面（.oe-page 已给 page class），正文沿用 jdoc 的 .docx-body 样式 */
.oe-view-paged .oe-page table { border-collapse: collapse; }
.oe-view-paged .oe-page img { max-width: 100%; }
/* jdoc 默认样式表给 .docx-page 加了外边距与阴影，分页容器自己有间距，去掉避免页间距翻倍 */
.oe-view-paged .oe-page.docx-page, .oe-page-edit { margin: 0; box-shadow: var(--oe-shadow); }
.oe-hint { padding: 16px; color: var(--oe-muted); }
.oe-error { padding: 16px; color: var(--oe-danger); white-space: pre-wrap; font-family: ui-monospace, monospace; }

/* ---------- DOCX 块编辑器（分页纸面内） ---------- */
/* 纸面恒白，故正文色固定为深色：--oe-text 在深色主题下是浅色，白底上不可读 */
.oe-doc .oe-page-body { font-size: 13px; color: #1f2328; }
.oe-doc .oe-page { background: #fff; }

/* ---------- DOCX 块编辑器 ---------- */
/* .oe-doc / .oe-pages 的纵向排布与外边距见上方「分页纸面」段 */
.oe-block { position: relative; margin: 2px 0; border-radius: 6px; border: 1px solid transparent; }
.oe-block:hover { border-color: var(--oe-border); }
.oe-block.sel { border-color: var(--oe-accent); background: var(--oe-accent-soft); }
.oe-block .oe-grip { position: absolute; left: -30px; top: 4px; display: none; flex-direction: column; gap: 2px; }
.oe-block:hover .oe-grip, .oe-block.sel .oe-grip { display: flex; }
.oe-grip button { appearance: none; border: 1px solid var(--oe-border); background: var(--oe-surface); border-radius: 5px; width: 24px; height: 22px; cursor: pointer; color: var(--oe-muted); display: inline-flex; align-items: center; justify-content: center; padding: 0; }
.oe-grip button:hover { color: var(--oe-accent); border-color: var(--oe-accent); }
.oe-grip svg { width: 13px; height: 13px; }
.oe-edit { outline: none; min-height: 1.5em; }
.oe-edit:focus { background: var(--vscode-editor-selectionBackground, #fffbe6); border-radius: 4px; }
.oe-edit[data-empty="1"]::before { content: attr(data-ph); color: #b6bcc7; }
.oe-block[data-type="heading"] .oe-edit { font-weight: 700; line-height: 1.3; margin: .5em 0 .3em; }
.oe-block[data-type="heading"][data-level="1"] .oe-edit { font-size: 24px; }
.oe-block[data-type="heading"][data-level="2"] .oe-edit { font-size: 19px; }
.oe-block[data-type="heading"][data-level="3"] .oe-edit { font-size: 16px; }
.oe-block[data-type="heading"][data-level="4"] .oe-edit { font-size: 14px; }
.oe-block[data-type="heading"][data-level="5"] .oe-edit { font-size: 13px; }
.oe-block[data-type="heading"][data-level="6"] .oe-edit { font-size: 12.5px; }
/* 章节号：解析自 numbering 的只读装饰，与正文留出间距 */
.oe-h-num, .oe-edit .oe-h-num { margin-right: .45em; white-space: nowrap; user-select: none; }
.oe-edit .oe-h-num { display: inline-block; }
.oe-block[data-type="paragraph"] .oe-edit { margin: .35em 0; }
.oe-block[data-type="quote"] .oe-edit { border-left: 3px solid var(--oe-accent); padding-left: 12px; color: var(--oe-muted); font-style: italic; }
.oe-block[data-type="code"] .oe-edit { font-family: var(--vscode-editor-font-family, ui-monospace, monospace); background: #f6f8fa; padding: 10px 12px; border-radius: 6px; white-space: pre-wrap; }
.oe-block[data-type="hr"] { padding: 6px 0; }
.oe-block[data-type="hr"] hr { border: 0; border-top: 2px solid var(--oe-border); margin: 0; }
.oe-block[data-type="image"] img { max-width: 100%; border-radius: 4px; }
.oe-block[data-type="image"] .cap { color: var(--oe-muted); font-size: 12px; text-align: center; margin-top: 4px; }
.oe-block[data-type="pageBreak"] { text-align: center; color: #b6bcc7; font-size: 12px; border-top: 2px dashed var(--oe-border); padding: 8px 0; }
.oe-block[data-type="list"] .oe-li { display: flex; gap: 8px; align-items: flex-start; margin: 2px 0; }
.oe-block[data-type="list"] .oe-li .mk { color: var(--oe-muted); min-width: 18px; text-align: right; user-select: none; padding-top: 2px; }
.oe-block[data-type="list"] .oe-li .oe-edit { flex: 1; }
.oe-doc table.oe-tbl { border-collapse: collapse; width: 100%; margin: 8px 0; }
.oe-doc table.oe-tbl td, .oe-doc table.oe-tbl th { border: 1px solid var(--oe-border); padding: 6px 8px; vertical-align: top; min-width: 40px; }
.oe-doc table.oe-tbl th { background: #f7f8fa; font-weight: 600; }
.oe-doc table.oe-tbl td .oe-edit, .oe-doc table.oe-tbl th .oe-edit { min-height: 1.4em; }
/* 目录块（编辑模式只读还原） */
.oe-block[data-type="toc"] .oe-toc { display: flex; flex-direction: column; gap: 1px; padding: 2px 0 6px; }
.oe-toc-row { display: flex; align-items: baseline; gap: .3em; line-height: 1.5; }
.oe-toc-row[data-level="1"] { font-weight: 600; }
.oe-toc-row[data-level="2"] { padding-left: 1.6em; }
.oe-toc-row[data-level="3"] { padding-left: 3.2em; font-size: .95em; }
.oe-toc-text { white-space: pre; }
.oe-toc-dots { flex: 1 1 auto; border-bottom: 1px dotted currentColor; opacity: .45; transform: translateY(-.28em); min-width: 1.5em; }
.oe-toc-page { white-space: nowrap; }

/* 插入条：零高度 + 浮层。分页时它不能占版面，否则每个块都凭空多出约 30px，
   页数会明显虚高（上百个块的文档能多出好几页）。 */
.oe-insert { position: relative; z-index: 5; height: 0; overflow: visible; display: flex; gap: 6px; justify-content: center; opacity: 0; transition: opacity .12s; }
.oe-doc:hover .oe-insert { opacity: .55; }
.oe-insert:hover { opacity: 1; }
.oe-insert button { appearance: none; border: 1px dashed var(--oe-border); background: var(--oe-surface); color: var(--oe-muted); border-radius: 6px; padding: 2px 10px; font-size: 11px; line-height: 1.6; cursor: pointer; }
.oe-insert button:hover { color: var(--oe-accent); border-color: var(--oe-accent); }

/* ---------- XLSX 网格编辑器 ---------- */
.oe-xl { display: flex; flex-direction: column; height: 100%; }
.oe-xl-tabs { display: flex; gap: 2px; padding: 6px 12px 0; background: var(--oe-surface); border-bottom: 1px solid var(--oe-border); overflow-x: auto; }
.oe-xl-tabs button { appearance: none; border: 1px solid var(--oe-border); border-bottom: 0; background: var(--oe-toolbar); color: var(--oe-muted); padding: 5px 14px; font-size: 13px; cursor: pointer; border-radius: 6px 6px 0 0; white-space: nowrap; }
.oe-xl-tabs button.active { background: var(--oe-surface); color: var(--oe-accent); font-weight: 600; }
.oe-xl-tabs .add { color: var(--oe-accent); font-weight: 700; }
.oe-grid-wrap { flex: 1; overflow: auto; background: var(--oe-surface); }
.oe-grid { border-collapse: separate; border-spacing: 0; font-size: 13px; table-layout: fixed; }
.oe-grid th, .oe-grid td { border-right: 1px solid var(--oe-border); border-bottom: 1px solid var(--oe-border); height: 26px; padding: 0; position: relative; }
.oe-grid thead th { background: #f7f8fa; color: var(--oe-muted); font-weight: 600; text-align: center; position: sticky; top: 0; z-index: 3; user-select: none; }
.oe-grid tbody th.oe-rownum { background: #f7f8fa; color: var(--oe-muted); font-weight: 500; text-align: center; position: sticky; left: 0; z-index: 2; user-select: none; width: 48px; }
.oe-grid thead th.oe-corner { left: 0; z-index: 4; }
.oe-grid td .cell { width: 100%; height: 100%; min-height: 26px; padding: 3px 6px; outline: none; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; cursor: cell; }
.oe-grid td.inrange { background: var(--oe-accent-soft); }
.oe-grid td.active { outline: 2px solid var(--oe-accent); outline-offset: -2px; }
.oe-grid td .cell[contenteditable="true"] { background: #fff; color: #1f2328; white-space: normal; overflow: auto; cursor: text; box-shadow: inset 0 0 0 2px var(--oe-accent); }
.oe-grid .col-res { position: absolute; top: 0; right: -3px; width: 6px; height: 100%; cursor: col-resize; z-index: 1; }
.oe-grid .fz { position: sticky; z-index: 2; }
.oe-grid thead th.fz { z-index: 6; }
.oe-grid tbody th.oe-rownum.fz { z-index: 4; }

/* ---------- toast ---------- */
.oe-toast { position: fixed; left: 50%; bottom: 28px; transform: translateX(-50%); background: var(--vscode-notifications-background, #1f2430); color: var(--vscode-notifications-foreground, #fff); border: 1px solid var(--jc-border); padding: 8px 16px; border-radius: 8px; font-size: 13px; box-shadow: var(--oe-shadow); opacity: 0; pointer-events: none; transition: opacity .15s, transform .15s; z-index: 50; max-width: 80%; }
.oe-toast.show { opacity: .96; }
.oe-toast.err { background: var(--oe-danger); }

/* ---------- 深色主题（VS Code 注入 data-oe-theme，或跟随系统偏好） ---------- */
[data-oe-theme="dark"] {
  --oe-bg:#1b1d23; --oe-surface:#23262d; --oe-border:#343841; --oe-text:#e6e8ec; --oe-muted:#9aa1ac;
  --oe-accent:#5b9cf5; --oe-accent-soft:#2a3b52; --oe-toolbar:#1f2228; --oe-danger:#e07a7a;
  --oe-shadow:0 1px 2px rgba(0,0,0,.5), 0 1px 3px rgba(0,0,0,.4);
}
[data-oe-theme="dark"] .oe-edit:focus { background:#2c3340; }
[data-oe-theme="dark"] .oe-doc table.oe-tbl th { background:#2a2e36; }
/* 纸面恒白（见 .oe-page 注释）：纸面内的深色主题底色/文字一律换回浅色，
   否则白纸上会出现暗底暗字、深底深字。XLSX 网格是整幅 UI（不是纸面），保持跟随主题。 */
[data-oe-theme="dark"] .oe-page .oe-edit:focus { background:#fffbe6; }
[data-oe-theme="dark"] .oe-page table.oe-tbl th { background:#f7f8fa; }
[data-oe-theme="dark"] .oe-page table.oe-tbl td,
[data-oe-theme="dark"] .oe-page table.oe-tbl th { border-color:#d8dce3; }
[data-oe-theme="dark"] .oe-page .oe-block:hover { border-color:#d8dce3; }
[data-oe-theme="dark"] .oe-page .oe-block.sel { background:#dbeafe; border-color:#2b6cb0; }
[data-oe-theme="dark"] .oe-page .oe-insert button { color:#6b7280; border-color:#c9ced6; background:#fff; }
[data-oe-theme="dark"] .oe-page .oe-insert button:hover { color:#2b6cb0; border-color:#2b6cb0; }
[data-oe-theme="dark"] .oe-page .oe-page-num { color:#8a919c; }
[data-oe-theme="dark"] .oe-grid thead th,
[data-oe-theme="dark"] .oe-grid tbody th.oe-rownum { background:#2a2e36; }
[data-oe-theme="dark"] .oe-grid td .cell[contenteditable="true"] { background:#1f2430; }
@media (prefers-color-scheme: dark) {
  :root:not([data-oe-theme="light"]) {
    --oe-bg:#1b1d23; --oe-surface:#23262d; --oe-border:#343841; --oe-text:#e6e8ec; --oe-muted:#9aa1ac;
    --oe-accent:#5b9cf5; --oe-accent-soft:#2a3b52; --oe-toolbar:#1f2228; --oe-danger:#e07a7a;
    --oe-shadow:0 1px 2px rgba(0,0,0,.5), 0 1px 3px rgba(0,0,0,.4);
  }
}
`;
