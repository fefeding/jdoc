// office-editor 样式：外壳（工具栏/内容/状态栏）+ 块式 DOCX + 网格 XLSX。

export const officeCss = `
:root {
  --oe-bg: #f3f4f6;
  --oe-surface: #ffffff;
  --oe-border: #e2e5ea;
  --oe-text: #1f2430;
  --oe-muted: #6b7280;
  --oe-accent: #2b6cb0;
  --oe-accent-soft: #e8f0fb;
  --oe-danger: #d64545;
  --oe-toolbar: #fafbfc;
  --oe-shadow: 0 1px 2px rgba(16,24,40,.06), 0 1px 3px rgba(16,24,40,.1);
}
* { box-sizing: border-box; }
html, body, #app { height: 100%; margin: 0; }
body {
  font: 14px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  color: var(--oe-text);
  background: var(--oe-bg);
}
.oe { display: flex; flex-direction: column; height: 100%; }
.oe-top { display: flex; align-items: center; gap: 8px; padding: 6px 12px; background: var(--oe-surface); border-bottom: 1px solid var(--oe-border); }
.oe-top .oe-title { font-weight: 600; font-size: 13px; color: var(--oe-text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 40%; }
.oe-top .oe-title .dirty { color: var(--oe-accent); }
.oe-seg { display: inline-flex; border: 1px solid var(--oe-border); border-radius: 8px; overflow: hidden; }
.oe-seg button { appearance: none; border: 0; background: var(--oe-surface); color: var(--oe-muted); padding: 5px 12px; font-size: 13px; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; }
.oe-seg button + button { border-left: 1px solid var(--oe-border); }
.oe-seg button.active { background: var(--oe-accent); color: #fff; }
.oe-seg svg { width: 15px; height: 15px; }
.oe-spacer { flex: 1; }

.oe-toolbar { display: flex; flex-wrap: wrap; align-items: center; gap: 4px; padding: 6px 12px; background: var(--oe-toolbar); border-bottom: 1px solid var(--oe-border); min-height: 42px; }
.oe-toolbar:empty { display: none; }
.oe-tb { appearance: none; border: 1px solid transparent; background: transparent; color: var(--oe-text); height: 30px; min-width: 30px; padding: 0 7px; border-radius: 6px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 5px; font-size: 13px; }
.oe-tb:hover { background: var(--oe-accent-soft); }
.oe-tb.active { background: var(--oe-accent); color: #fff; border-color: var(--oe-accent); }
.oe-tb:disabled { opacity: .4; cursor: default; background: transparent; }
.oe-tb svg { width: 17px; height: 17px; }
.oe-tb.oe-glyph { font-weight: 700; }
.oe-sep { width: 1px; align-self: stretch; margin: 4px 4px; background: var(--oe-border); }
.oe-select { height: 30px; border: 1px solid var(--oe-border); border-radius: 6px; background: var(--oe-surface); color: var(--oe-text); padding: 0 6px; font-size: 13px; cursor: pointer; }
.oe-color { width: 30px; height: 30px; padding: 0; border: 1px solid var(--oe-border); border-radius: 6px; background: var(--oe-surface); cursor: pointer; }

.oe-content { flex: 1; overflow: auto; position: relative; }
.oe-statusbar { display: flex; align-items: center; gap: 12px; padding: 4px 12px; background: var(--oe-surface); border-top: 1px solid var(--oe-border); font-size: 12px; color: var(--oe-muted); }
.oe-statusbar .grow { flex: 1; }

/* ---------- 查看模式（DOCX/XLSX 只读预览） ---------- */
.oe-view { padding: 24px; display: flex; justify-content: center; }
.oe-view .page { background: var(--oe-surface); box-shadow: var(--oe-shadow); border-radius: 4px; padding: 48px 56px; max-width: 820px; width: 100%; }
.oe-view .page table { border-collapse: collapse; }
.oe-view img { max-width: 100%; }
.oe-hint { padding: 16px; color: var(--oe-muted); }
.oe-error { padding: 16px; color: var(--oe-danger); white-space: pre-wrap; font-family: ui-monospace, monospace; }

/* ---------- DOCX 块编辑器 ---------- */
.oe-doc { display: flex; flex-direction: column; align-items: center; padding: 24px 16px 80px; }
.oe-doc .sheet { background: var(--oe-surface); box-shadow: var(--oe-shadow); border-radius: 4px; width: 100%; max-width: 820px; padding: 40px 52px; }
.oe-block { position: relative; margin: 2px 0; border-radius: 6px; border: 1px solid transparent; }
.oe-block:hover { border-color: var(--oe-border); }
.oe-block.sel { border-color: var(--oe-accent); background: var(--oe-accent-soft); }
.oe-block .oe-grip { position: absolute; left: -30px; top: 4px; display: none; flex-direction: column; gap: 2px; }
.oe-block:hover .oe-grip, .oe-block.sel .oe-grip { display: flex; }
.oe-grip button { appearance: none; border: 1px solid var(--oe-border); background: var(--oe-surface); border-radius: 5px; width: 24px; height: 22px; cursor: pointer; color: var(--oe-muted); display: inline-flex; align-items: center; justify-content: center; padding: 0; }
.oe-grip button:hover { color: var(--oe-accent); border-color: var(--oe-accent); }
.oe-grip svg { width: 13px; height: 13px; }
.oe-edit { outline: none; min-height: 1.5em; }
.oe-edit:focus { background: #fffbe6; border-radius: 4px; }
.oe-edit[data-empty="1"]::before { content: attr(data-ph); color: #b6bcc7; }
.oe-block[data-type="heading"] .oe-edit { font-weight: 700; line-height: 1.3; margin: .5em 0 .3em; }
.oe-block[data-type="heading"][data-level="1"] .oe-edit { font-size: 28px; }
.oe-block[data-type="heading"][data-level="2"] .oe-edit { font-size: 23px; }
.oe-block[data-type="heading"][data-level="3"] .oe-edit { font-size: 19px; }
.oe-block[data-type="heading"][data-level="4"] .oe-edit { font-size: 16px; }
.oe-block[data-type="heading"][data-level="5"] .oe-edit { font-size: 15px; }
.oe-block[data-type="heading"][data-level="6"] .oe-edit { font-size: 14px; }
.oe-block[data-type="paragraph"] .oe-edit { margin: .35em 0; }
.oe-block[data-type="quote"] .oe-edit { border-left: 3px solid var(--oe-accent); padding-left: 12px; color: var(--oe-muted); font-style: italic; }
.oe-block[data-type="code"] .oe-edit { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background: #f6f8fa; padding: 10px 12px; border-radius: 6px; white-space: pre-wrap; }
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
.oe-insert { align-self: stretch; max-width: 820px; width: 100%; display: flex; gap: 6px; justify-content: center; padding: 4px 0; opacity: 0; transition: opacity .12s; }
.oe-doc:hover .oe-insert { opacity: .55; }
.oe-insert:hover { opacity: 1; }
.oe-insert button { appearance: none; border: 1px dashed var(--oe-border); background: var(--oe-surface); color: var(--oe-muted); border-radius: 6px; padding: 3px 10px; font-size: 12px; cursor: pointer; }
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
.oe-grid td .cell[contenteditable="true"] { background: #fff; white-space: normal; overflow: auto; cursor: text; box-shadow: inset 0 0 0 2px var(--oe-accent); }
.oe-grid .col-res { position: absolute; top: 0; right: -3px; width: 6px; height: 100%; cursor: col-resize; z-index: 1; }
.oe-grid .fz { position: sticky; z-index: 2; }
.oe-grid thead th.fz { z-index: 6; }
.oe-grid tbody th.oe-rownum.fz { z-index: 4; }

/* ---------- toast ---------- */
.oe-toast { position: fixed; left: 50%; bottom: 28px; transform: translateX(-50%); background: #1f2430; color: #fff; padding: 8px 16px; border-radius: 8px; font-size: 13px; box-shadow: var(--oe-shadow); opacity: 0; pointer-events: none; transition: opacity .15s, transform .15s; z-index: 50; max-width: 80%; }
.oe-toast.show { opacity: .96; }
.oe-toast.err { background: var(--oe-danger); }
`;
