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
  /* 选中态前景：必须与 --jc-active 配对使用。
     不要拿 --jc-accent 当选中文字色压在 --jc-accent-soft 上——两者分别来自
     button-background 与 badge-background，VS Code 不保证它们之间有对比度
     （实测同一深色主题下可低至 1.3:1，字几乎看不见）。 */
  --jc-active-fg: var(--vscode-list-activeSelectionForeground, #e6e9ef);
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
  --jc-active-fg: var(--vscode-list-activeSelectionForeground, #1f2430);
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

/* ---------- 缩放簇（zoom.ts，所有格式共用） ---------- */
.jc-zoom { display: inline-flex; align-items: center; gap: 2px; flex: 0 0 auto; }
.jc-zoom .chip { appearance: none; border: 0; background: transparent; color: inherit; font: inherit; padding: 2px 7px; border-radius: 4px; cursor: pointer; min-width: 22px; line-height: 1.4; }
.jc-zoom .chip:hover { background: var(--jc-hover); }
.jc-zoom .chip.on { background: var(--jc-active); color: var(--jc-text); }

/* 状态栏里的普通 chip（转换格式按钮）：与缩放簇按钮同一套视觉，
   否则会渲染成突兀的原生灰色按钮，在深色状态栏上几乎看不出来。 */
.jc-statusbar .chip {
  appearance: none; border: 0; background: transparent; color: inherit; font: inherit;
  padding: 2px 7px; border-radius: 4px; cursor: pointer; line-height: 1.4;
}
.jc-statusbar .chip:hover { background: var(--jc-hover); }

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
/**
 * 向宿主发消息的函数（由各 webview 入口注入）
 *
 * **不能在这里自己调 acquireVsCodeApi()**：VS Code 规定该函数每个 webview 只能调用
 * 一次，第二次会抛「An instance of the VS Code API has already been acquired」。
 * 各入口（main.ts / office-editor/main.ts）都在模块顶层已经调过一次，这里再调就会
 * 在点击时抛异常——表现为「按钮点了没反应」，而 TAB 右键（走 VS Code 命令，不碰
 * webview API）却正常，很难定位。故统一由入口注入自己的 post 函数。
 */
import { t } from './i18n';

let hostPost: (msg: unknown) => void = () => undefined;

/** 注入向宿主发消息的函数（入口在模块顶层调用一次即可） */
export function setHostPostMessage(fn: (msg: unknown) => void): void {
    hostPost = fn;
}

/**
 * 状态栏「转换格式」按钮
 *
 * 本扩展打开文档用的是 custom editor，正文渲染在 webview 里：内容区的右键事件被
 * webview 吞掉，VS Code 的 editor/context 菜单**根本不会出现**——只靠右键菜单，
 * 转换功能在编辑器里等于没有入口（这正是之前「移除 editor/title、全部收敛到右键」
 * 那次改动留下的死角）。这里在常驻的状态栏上给出可靠入口，点击后交回宿主执行
 * jdoc.convert；宿主侧 protocol 已加 `{ type: 'convert' }` 消息与之配对。
 */
export function createConvertChip(): HTMLButtonElement {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'chip';
    btn.textContent = t('chrome.convert');
    btn.title = t('chrome.convertTitle');
    btn.addEventListener('click', () => postConvert());
    return btn;
}

/** 当前文档的 uri（init 时由宿主下发） */
let currentDocUri: string | null = null;

/** 记录当前文档 uri：转换请求必须带上它，否则宿主只能靠 activeTextEditor 猜，
 *  在同时开着其它文本文件时会转错对象（Markdown 走内置编辑器时尤其明显）。 */
export function setCurrentDocUri(uri: string | null): void {
    currentDocUri = uri;
}

/**
 * 转换请求的状态反馈通道（各入口注入自己的状态栏）
 *
 * 没有它，转换链路上任何一环出问题都表现为「按钮点了没反应」——消息没送达、
 * 选择框没弹出、目标格式不支持，用户无从判断卡在哪一步。
 */
let convertStatusSink: ((text: string) => void) | null = null;

/** 注入转换状态显示通道（一般指向状态栏文字） */
export function setConvertStatusSink(fn: ((text: string) => void) | null): void {
    convertStatusSink = fn;
}

/** 宿主已回执：目标格式选择框即将弹出 */
export function notifyConvertAck(): void {
    convertStatusSink?.(t('chrome.convertAck'));
}

/** 向宿主请求转换当前文档 */
export function postConvert(): void {
    convertStatusSink?.(t('chrome.convertRequesting'));
    hostPost({ type: 'convert', uri: currentDocUri ?? undefined });
}

/**
 * 「转换格式」右键菜单项
 *
 * 内容区右键时与状态栏按钮走同一条路径（postConvert），保证两处行为一致。
 */
export function convertMenuEntry(): ContextMenuEntry {
    return {
        label: t('chrome.convertMenu'),
        action: () => postConvert()
    };
}

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

/* ---------- 文档内容区右键菜单 ---------- */

/** 右键菜单项 */
export interface ContextMenuEntry {
    label: string;
    /** 省略表示分隔线 */
    action?: () => void;
    /** 置灰不可用 */
    disabled?: boolean;
    /** 快捷键提示，如 "⌘C" */
    hint?: string;
}

const MENU_CSS = `
.jc-ctxmenu {
  position: fixed; z-index: 9999; min-width: 190px; padding: 4px 0;
  background: var(--vscode-menu-background, var(--jc-surface-2, #25282e));
  color: var(--vscode-menu-foreground, var(--jc-text, #e6e6e6));
  border: 1px solid var(--vscode-menu-border, var(--jc-border, rgba(255,255,255,.12)));
  border-radius: 6px; box-shadow: 0 6px 22px rgba(0,0,0,.4);
  font-size: 13px; line-height: 1.4; user-select: none;
}
.jc-ctxmenu-item {
  display: flex; align-items: center; gap: 16px;
  padding: 5px 14px; cursor: pointer; white-space: nowrap;
}
.jc-ctxmenu-item:hover:not(.disabled) { background: var(--vscode-menu-selectionBackground, var(--jc-hover, rgba(255,255,255,.1))); color: var(--vscode-menu-selectionForeground, inherit); }
.jc-ctxmenu-item.disabled { opacity: .45; cursor: default; }
.jc-ctxmenu-item .label { flex: 1 1 auto; }
.jc-ctxmenu-item .hint { opacity: .6; font-size: 12px; }
.jc-ctxmenu-sep { height: 1px; margin: 4px 0; background: var(--vscode-menu-separatorBackground, var(--jc-border, rgba(255,255,255,.12))); }
`;

/** 菜单是否已注入过样式 */
let menuCssInjected = false;

function ensureMenuCss(): void {
    if (menuCssInjected) return;
    menuCssInjected = true;
    const style = document.createElement('style');
    style.textContent = MENU_CSS;
    document.head.appendChild(style);
}

/** 当前打开的菜单（同一时刻只允许一个） */
let openMenu: HTMLElement | null = null;

function closeMenu(): void {
    openMenu?.remove();
    openMenu = null;
}

function showMenu(x: number, y: number, entries: ContextMenuEntry[]): void {
    ensureMenuCss();
    closeMenu();
    const menu = document.createElement('div');
    menu.className = 'jc-ctxmenu';
    for (const entry of entries) {
        if (!entry.action) {
            const sep = document.createElement('div');
            sep.className = 'jc-ctxmenu-sep';
            menu.appendChild(sep);
            continue;
        }
        const item = document.createElement('div');
        item.className = 'jc-ctxmenu-item' + (entry.disabled ? ' disabled' : '');
        const label = document.createElement('span');
        label.className = 'label';
        label.textContent = entry.label;
        item.appendChild(label);
        if (entry.hint) {
            const hint = document.createElement('span');
            hint.className = 'hint';
            hint.textContent = entry.hint;
            item.appendChild(hint);
        }
        if (!entry.disabled) {
            item.addEventListener('click', () => {
                closeMenu();
                entry.action?.();
            });
        }
        menu.appendChild(item);
    }
    document.body.appendChild(menu);
    openMenu = menu;

    // 贴边翻转：靠近右/下边缘时向内收，避免菜单被视口裁掉
    const r = menu.getBoundingClientRect();
    const left = x + r.width > window.innerWidth ? Math.max(4, x - r.width) : x;
    const top = y + r.height > window.innerHeight ? Math.max(4, y - r.height) : y;
    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;

    // 点击别处 / 滚动 / 失焦 / Esc 关闭
    const onDocDown = (e: MouseEvent): void => {
        if (!menu.contains(e.target as Node)) closeMenu();
    };
    const onKey = (e: KeyboardEvent): void => {
        if (e.key === 'Escape') closeMenu();
    };
    // 用 capture + once 组合：首次外部点击后解绑自身
    setTimeout(() => {
        document.addEventListener('mousedown', onDocDown, { capture: true });
        document.addEventListener('keydown', onKey, { capture: true });
        window.addEventListener('blur', closeMenu);
        window.addEventListener('resize', closeMenu);
    }, 0);
    const cleanup = (): void => {
        document.removeEventListener('mousedown', onDocDown, { capture: true });
        document.removeEventListener('keydown', onKey, { capture: true });
        window.removeEventListener('blur', closeMenu);
        window.removeEventListener('resize', closeMenu);
    };
    menu.addEventListener('remove', cleanup);
}

/**
 * 给文档内容区装上自定义右键菜单
 *
 * 本扩展用 custom editor 打开文档，正文渲染在 webview（iframe）里：内容区的右键
 * 事件被 webview 吞掉，**VS Code 的 editor/context 菜单根本不会出现**，用户只能看到
 * 浏览器默认的「复制/粘贴」。这里自绘菜单，把「转换格式」等文档级操作补进内容区右键，
 * 三条 webview 链路（office-editor / 顶层 webview / pptx-editor）共用这一份实现。
 *
 * @param getEntries 按当前状态生成菜单项（可依据选区/编辑模式动态变化）
 * @param isEditable 命中可编辑区域时是否保留浏览器原生菜单（让复制粘贴照常工作）
 */
export function installContextMenu(getEntries: () => ContextMenuEntry[], isEditable?: (target: HTMLElement) => boolean): void {
    document.addEventListener(
        'contextmenu',
        (e: MouseEvent) => {
            const target = e.target as HTMLElement | null;
            // 输入框 / 可编辑区域交给浏览器原生菜单（复制粘贴、撤销等原生行为更完整）
            if (target && isEditable?.(target)) return;
            e.preventDefault();
            showMenu(e.clientX, e.clientY, getEntries());
        },
        true
    );
    // 菜单打开时按 Esc 关闭（已监听），这里额外保证滚动时收起
    document.addEventListener('scroll', closeMenu, true);
}

/** 选区是否有文本（决定「复制」是否可用） */
export function hasTextSelection(): boolean {
    const sel = window.getSelection();
    return !!(sel && !sel.isCollapsed && sel.toString().length > 0);
}

/** 复制当前选区 */
export function copySelection(): void {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) return;
    const text = sel.toString();
    void navigator.clipboard?.writeText(text).catch(() => undefined);
}
