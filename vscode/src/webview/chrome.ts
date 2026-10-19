// 统一外观层：所有 webview 入口（查看器 / office-editor / pptx-editor）共用同一套
// 设计令牌、字体与底部状态栏。
//
// 顶部区域刻意不画任何外壳：文档标题、面包屑、未保存点都由 VS Code 编辑器标题栏/标签栏
// 呈现，webview 再画一份只会重复；查看/编辑切换与文档切换改为注册成 editor/title 命令，
// 由 host 通过 setMode 消息驱动（见 package.json 的 menus 与 extension.ts）。
//
// 主题：颜色全部取自 VS Code 注入 webview 的 `--vscode-*` 主题变量，用户切换颜色主题
// （含 High Contrast）时实时跟随，不需要自己维护深浅色两套配色。
// 兜底值按 `body.vscode-light / vscode-dark / vscode-high-contrast` 三类给出
// （这三个类由 VS Code 自动加到 webview 的 body 上）。

/** 共享外观样式表（设计令牌 + 字体 + 内容区 + 状态栏） */
export const CHROME_CSS = `
/* ---------- 设计令牌：优先跟随 VS Code 主题 ---------- */
body {
  --jc-font: var(--vscode-font-family, -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif);
  --jc-font-mono: var(--vscode-editor-font-family, ui-monospace, SFMono-Regular, Menlo, Consolas, 'Liberation Mono', monospace);

  --jc-bg: var(--vscode-editor-background, #1b1f27);
  --jc-surface: var(--vscode-editor-background, #232936);
  --jc-surface-2: var(--vscode-editorWidget-background, var(--vscode-sideBar-background, #1f242f));
  --jc-statusbar: var(--vscode-statusBar-background, var(--vscode-editorGroupHeader-tabsBackground, #232936));
  --jc-statusbar-fg: var(--vscode-statusBar-foreground, var(--vscode-descriptionForeground, #9aa3b2));
  --jc-border: var(--vscode-panel-border, #333a49);
  --jc-border-soft: var(--vscode-editorWidget-border, var(--vscode-panel-border, #333a49));
  --jc-text: var(--vscode-foreground, #e6e9ef);
  --jc-muted: var(--vscode-descriptionForeground, #9aa3b2);
  --jc-accent: var(--vscode-button-background, #4c8fe0);
  --jc-accent-fg: var(--vscode-button-foreground, #ffffff);
  --jc-accent-hover: var(--vscode-button-hoverBackground, #5f9de6);
  --jc-accent-soft: var(--vscode-badge-background, #2b3a55);
  --jc-accent-soft-fg: var(--vscode-badge-foreground, var(--vscode-textLink-foreground, #9ecbff));
  --jc-hover: var(--vscode-list-hoverBackground, rgba(255,255,255,.06));
  --jc-active: var(--vscode-list-activeSelectionBackground, var(--vscode-toolbar-activeBackground, rgba(76,143,224,.24)));
  --jc-focus: var(--vscode-focusBorder, #4c8fe0);
  --jc-link: var(--vscode-textLink-foreground, #4daafc);
  --jc-code-bg: var(--vscode-textCodeBlock-background, rgba(255,255,255,.07));
  --jc-code-fg: var(--vscode-textPreformat-foreground, var(--vscode-foreground, #e6e9ef));
  color-scheme: dark;
}
/* 浅色宿主：仅换兜底值，宿主变量存在时仍以宿主为准 */
body.vscode-light {
  --jc-bg: var(--vscode-editor-background, #ffffff);
  --jc-surface: var(--vscode-editor-background, #ffffff);
  --jc-surface-2: var(--vscode-editorWidget-background, var(--vscode-sideBar-background, #f8f8f8));
  --jc-statusbar: var(--vscode-statusBar-background, var(--vscode-editorGroupHeader-tabsBackground, #f8f8f8));
  --jc-statusbar-fg: var(--vscode-statusBar-foreground, var(--vscode-descriptionForeground, #616161));
  --jc-border: var(--vscode-panel-border, #e5e5e5);
  --jc-border-soft: var(--vscode-editorWidget-border, var(--vscode-panel-border, #e5e5e5));
  --jc-text: var(--vscode-foreground, #3b3b3b);
  --jc-muted: var(--vscode-descriptionForeground, #616161);
  --jc-accent: var(--vscode-button-background, #005fb8);
  --jc-accent-fg: var(--vscode-button-foreground, #ffffff);
  --jc-accent-hover: var(--vscode-button-hoverBackground, #0258a8);
  --jc-accent-soft: var(--vscode-badge-background, #c4c4c4);
  --jc-accent-soft-fg: var(--vscode-badge-foreground, #ffffff);
  --jc-hover: var(--vscode-list-hoverBackground, #e8e8e8);
  --jc-active: var(--vscode-list-activeSelectionBackground, #e8e8e8);
  --jc-focus: var(--vscode-focusBorder, #005fb8);
  --jc-link: var(--vscode-textLink-foreground, #005fb8);
  --jc-code-bg: var(--vscode-textCodeBlock-background, rgba(0,0,0,.05));
  color-scheme: light;
}
/* High Contrast：仍走同一套 vscode 变量，仅保证边框可见 */
body.vscode-high-contrast {
  --jc-border: var(--vscode-panel-border, #6fc3df);
  --jc-border-soft: var(--vscode-panel-border, #6fc3df);
  color-scheme: dark;
}

/* ---------- 字体与基础 ---------- */
body { font-family: var(--jc-font); }
button, input, select, textarea { font-family: inherit; }
:focus-visible { outline: 1px solid var(--jc-focus); outline-offset: -1px; }

/* ---------- 内容与状态栏 ---------- */
.jc-content { flex: 1 1 auto; overflow: auto; position: relative; background: var(--jc-bg); }
.jc-statusbar {
  display: flex; align-items: center; gap: 12px; flex: 0 0 auto;
  padding: 4px 12px; background: var(--jc-statusbar); border-top: 1px solid var(--jc-border);
  font-size: 12px; color: var(--jc-statusbar-fg);
}
.jc-statusbar .jc-grow { flex: 1 1 auto; }
.jc-statusbar .jc-status-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

/* ---------- 滚动条（跟随 VS Code） ---------- */
.jc-content::-webkit-scrollbar, .oe-content::-webkit-scrollbar, .jdoc-md ::-webkit-scrollbar { width: 10px; height: 10px; }
.jc-content::-webkit-scrollbar-thumb, .oe-content::-webkit-scrollbar-thumb, .jdoc-md ::-webkit-scrollbar-thumb {
  background: var(--vscode-scrollbarSlider-background, rgba(121,121,121,.4)); border-radius: 6px;
  border: 2px solid transparent; background-clip: content-box;
}
.jc-content::-webkit-scrollbar-thumb:hover, .oe-content::-webkit-scrollbar-thumb:hover, .jdoc-md ::-webkit-scrollbar-thumb:hover {
  background: var(--vscode-scrollbarSlider-hoverBackground, rgba(100,100,100,.7)); border: 2px solid transparent; background-clip: content-box;
}
.jc-content::-webkit-scrollbar-thumb:active, .oe-content::-webkit-scrollbar-thumb:active, .jdoc-md ::-webkit-scrollbar-thumb:active {
  background: var(--vscode-scrollbarSlider-activeBackground, rgba(85,85,85,.8)); border: 2px solid transparent; background-clip: content-box;
}
`;
/** 创建统一底部状态栏 */
export function createStatusBar(): { el: HTMLElement; textEl: HTMLElement } {
    const textEl = document.createElement('span');
    textEl.className = 'jc-status-text';
    const grow = document.createElement('span');
    grow.className = 'jc-grow';
    const el = document.createElement('div');
    el.className = 'jc-statusbar';
    el.append(textEl, grow);
    return { el, textEl };
}
