// 统一外观层：所有 webview 入口（查看器 / office-editor / pptx-editor）共用同一套
// 顶栏布局——「文档标题 + 格式徽标在左，查看/编辑分段切换在右上，状态栏在底部」。
//
// 主题：extension 在 <html data-oe-theme="dark|light"> 上注入主题，深浅色由此驱动；
// 未注入时跟随系统 prefers-color-scheme。

/** 共享外观样式表（设计令牌 + 顶栏 + 分段切换 + 状态栏） */
export const CHROME_CSS = `
:root {
  --jc-bg: #f3f4f6;
  --jc-surface: #ffffff;
  --jc-surface-2: #fafbfc;
  --jc-border: #e2e5ea;
  --jc-text: #1f2430;
  --jc-muted: #6b7280;
  --jc-accent: #2b6cb0;
  --jc-accent-soft: #e8f0fb;
  --jc-shadow: 0 1px 2px rgba(16,24,40,.05);
}
html[data-oe-theme="dark"] {
  --jc-bg: #1b1f27;
  --jc-surface: #232936;
  --jc-surface-2: #1f242f;
  --jc-border: #333a49;
  --jc-text: #e6e9ef;
  --jc-muted: #9aa3b2;
  --jc-accent: #4c8fe0;
  --jc-accent-soft: #2b3a55;
  --jc-shadow: 0 1px 2px rgba(0,0,0,.3);
}
@media (prefers-color-scheme: dark) {
  html:not([data-oe-theme="light"]) {
    --jc-bg: #1b1f27; --jc-surface: #232936; --jc-surface-2: #1f242f; --jc-border: #333a49;
    --jc-text: #e6e9ef; --jc-muted: #9aa3b2; --jc-accent: #4c8fe0; --jc-accent-soft: #2b3a55;
  }
}

/* ---------- 顶栏 ---------- */
.jc-top {
  display: flex; align-items: center; gap: 10px; flex: 0 0 auto;
  padding: 8px 14px; background: var(--jc-surface);
  border-bottom: 1px solid var(--jc-border); box-shadow: var(--jc-shadow);
}
.jc-top-title { display: flex; align-items: center; gap: 8px; min-width: 0; font-size: 13px; font-weight: 600; color: var(--jc-text); }
.jc-top-title .jc-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.jc-badge {
  flex: 0 0 auto; font-size: 10px; font-weight: 700; letter-spacing: .6px;
  padding: 2px 7px; border-radius: 6px; text-transform: uppercase;
  background: var(--jc-accent-soft); color: var(--jc-accent);
}
.jc-dirty { width: 7px; height: 7px; border-radius: 50%; background: transparent; flex: 0 0 auto; }
.jc-dirty.on { background: var(--jc-accent); }
.jc-spacer { flex: 1 1 auto; }

/* ---------- 分段切换（查看/编辑） ---------- */
.jc-seg { display: inline-flex; gap: 2px; padding: 2px; border-radius: 9px; background: var(--jc-surface-2); border: 1px solid var(--jc-border); }
.jc-seg button {
  appearance: none; border: 0; background: transparent; color: var(--jc-muted);
  height: 26px; padding: 0 11px; border-radius: 7px; font-size: 12px; font-weight: 600;
  cursor: pointer; display: inline-flex; align-items: center; gap: 6px; transition: background .12s, color .12s;
}
.jc-seg button:hover { color: var(--jc-text); }
.jc-seg button.active { background: var(--jc-accent); color: #fff; box-shadow: 0 1px 2px rgba(16,24,40,.18); }
.jc-seg button svg { width: 14px; height: 14px; }

/* ---------- 文档切换标签（右上角） ---------- */
.jc-doctabs { max-width: 46vw; overflow-x: auto; overflow-y: hidden; scrollbar-width: none; margin-left: 6px; }
.jc-doctabs::-webkit-scrollbar { display: none; }
.jc-doctabs button { max-width: 190px; }
.jc-doctab-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.jc-doctab-dot {
  width: 6px; height: 6px; border-radius: 50%; flex: 0 0 auto;
  background: currentColor; opacity: 0;
}
.jc-doctab.dirty .jc-doctab-dot { opacity: .9; }

/* ---------- 内容与状态栏 ---------- */
.jc-content { flex: 1 1 auto; overflow: auto; position: relative; }
.jc-statusbar {
  display: flex; align-items: center; gap: 12px; flex: 0 0 auto;
  padding: 5px 14px; background: var(--jc-surface); border-top: 1px solid var(--jc-border);
  font-size: 12px; color: var(--jc-muted);
}
.jc-statusbar .jc-grow { flex: 1 1 auto; }
.jc-statusbar .jc-status-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

/* ---------- 滚动条 ---------- */
.jc-content::-webkit-scrollbar, .jdoc-md ::-webkit-scrollbar { width: 10px; height: 10px; }
.jc-content::-webkit-scrollbar-thumb, .jdoc-md ::-webkit-scrollbar-thumb { background: var(--jc-border); border-radius: 6px; border: 2px solid transparent; background-clip: content-box; }
.jc-content::-webkit-scrollbar-thumb:hover, .jdoc-md ::-webkit-scrollbar-thumb:hover { background: var(--jc-muted); border: 2px solid transparent; background-clip: content-box; }
`;

/** 分段切换的一项 */
export interface ChromeMode {
    /** 模式标识（调用方自定义语义） */
    id: string;
    /** 按钮文案 */
    label: string;
    /** 图标名（共享图标集，可选） */
    icon?: string;
}

/** createTopBar 的返回 */
export interface TopBar {
    el: HTMLElement;
    /** 标题文本节点容器 */
    nameEl: HTMLElement;
    /** 未保存标记 */
    dirtyEl: HTMLElement;
    /** 格式徽标 */
    badgeEl: HTMLElement;
    /** 切换高亮段 */
    setActive(id: string): void;
    /** 更新标题（含未保存点） */
    setTitle(text: string, dirty?: boolean): void;
    /** 更新格式徽标 */
    setBadge(text: string): void;
}

/** 创建统一顶栏：标题/徽标在左，分段切换在右上。 */
export function createTopBar(
    modes: ChromeMode[],
    current: string,
    onMode: (id: string) => void,
    icons: { node(name: string): Node }
): TopBar {
    const nameEl = document.createElement('span');
    nameEl.className = 'jc-name';
    nameEl.title = '';
    const dirtyEl = document.createElement('span');
    dirtyEl.className = 'jc-dirty';
    dirtyEl.title = '有未保存的修改';
    const badgeEl = document.createElement('span');
    badgeEl.className = 'jc-badge';

    const titleEl = document.createElement('div');
    titleEl.className = 'jc-top-title';
    titleEl.append(badgeEl, nameEl, dirtyEl);

    const buttons = modes.map((m) => {
        const b = document.createElement('button');
        b.type = 'button';
        b.dataset.mode = m.id;
        b.title = m.label;
        if (m.icon) b.append(icons.node(m.icon));
        b.append(document.createTextNode(m.label));
        b.onclick = () => onMode(m.id);
        return b;
    });
    const seg = document.createElement('div');
    seg.className = 'jc-seg';
    buttons.forEach((b) => seg.append(b));

    const spacer = document.createElement('span');
    spacer.className = 'jc-spacer';
    const el = document.createElement('div');
    el.className = 'jc-top';
    el.append(titleEl, spacer, seg);

    const setActive = (id: string) => {
        buttons.forEach((b) => b.classList.toggle('active', b.dataset.mode === id));
    };
    setActive(current);

    return {
        el,
        nameEl,
        dirtyEl,
        badgeEl,
        setActive,
        setTitle(text: string, dirty = false) {
            nameEl.textContent = text;
            nameEl.title = text + (dirty ? '（有未保存的修改）' : '');
            dirtyEl.classList.toggle('on', dirty);
        },
        setBadge(text: string) {
            badgeEl.textContent = text;
            badgeEl.style.display = text ? '' : 'none';
        }
    };
}

/** 已打开文档标签的数据项（与 host 的 DocTabInfo 对应，webview 侧不关心 format） */
export interface DocTabItem {
    uri: string;
    title: string;
    dirty?: boolean;
}

/** createDocTabs 的返回 */
export interface DocTabs {
    el: HTMLElement;
    /** 用 host 广播的文档列表刷新标签；只有一个文档时自动隐藏。 */
    update(items: DocTabItem[], activeUri: string): void;
}

/**
 * 顶栏右上角的「文档切换标签」：与查看/编辑分段切换同一套视觉，点击切换已打开的文档。
 * 只在打开 2 个及以上文档时显示，避免单文档时白占位置。
 */
export function createDocTabs(onSelect: (uri: string) => void): DocTabs {
    const el = document.createElement('div');
    el.className = 'jc-seg jc-doctabs';
    el.setAttribute('role', 'tablist');
    el.style.display = 'none';
    let signature = '';

    const update = (items: DocTabItem[], activeUri: string): void => {
        const sig =
            activeUri + '|' + items.map((i) => `${i.uri}:${i.title}:${i.dirty ? 1 : 0}`).join(',');
        if (sig === signature) return;
        signature = sig;
        el.innerHTML = '';
        el.style.display = items.length > 1 ? '' : 'none';
        for (const item of items) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className =
                'jc-doctab' +
                (item.uri === activeUri ? ' active' : '') +
                (item.dirty ? ' dirty' : '');
            btn.title = item.title;
            btn.setAttribute('role', 'tab');
            btn.setAttribute('aria-selected', item.uri === activeUri ? 'true' : 'false');
            const name = document.createElement('span');
            name.className = 'jc-doctab-name';
            name.textContent = item.title;
            const dot = document.createElement('span');
            dot.className = 'jc-doctab-dot';
            btn.append(name, dot);
            btn.onclick = () => onSelect(item.uri);
            el.append(btn);
        }
    };

    return { el, update };
}

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
