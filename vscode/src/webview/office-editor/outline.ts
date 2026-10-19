// 左侧目录 / 编号索引栏：把文档标题抽成一棵可点击跳转的大纲。
//
// 数据来源是渲染后的 DOM（而非 blocks）：查看模式与编辑模式的标题元素形态完全不同
// （h1~h6 vs contenteditable 块），统一从 DOM 上收集锚点，两种模式共用一份实现。

/** 大纲条目 */
export interface OutlineItem {
    /** 锚点 id（与正文元素的 data-oe-anchor 对应） */
    id: string;
    /** 标题文本（含章节号） */
    text: string;
    /** 层级 1-6 */
    level: number;
    /** 章节号前缀（"1." / "3.1."），无则空串 */
    number: string;
}

/** 正文标题元素上打锚点的属性 */
export const ANCHOR_ATTR = 'data-oe-anchor';

/** 收集轮廓（目录）时可选的选择器 */
const HEADING_SELECTOR = 'h1,h2,h3,h4,h5,h6,.oe-block[data-type="heading"] .oe-edit';

function headingLevelOf(el: HTMLElement): number {
    if (el.dataset.level) return Math.min(6, Math.max(1, Number(el.dataset.level) || 1));
    const tag = el.tagName.toLowerCase();
    const n = /^h([1-6])$/.exec(tag);
    return n ? Number(n[1]) : 1;
}

/**
 * 为容器内的标题打锚点并抽出大纲
 *
 * 已经打过锚点的（分页/重渲染后复用）直接沿用，避免 id 抖动导致侧栏高亮失效。
 */
export function buildOutline(root: HTMLElement): OutlineItem[] {
    const items: OutlineItem[] = [];
    const seen = new Map<HTMLElement, string>();
    let seq = 0;

    root.querySelectorAll<HTMLElement>(HEADING_SELECTOR).forEach((el) => {
        // 目录自身也会渲染出缩进行（带 .docx-toc-row），不要混进大纲
        if (el.closest('.docx-toc')) return;
        let id = el.getAttribute(ANCHOR_ATTR);
        if (!id) {
            id = 'h' + ++seq;
            el.setAttribute(ANCHOR_ATTR, id);
        }
        seen.set(el, id);
        const level = headingLevelOf(el);
        // 查看模式用 .docx-h-num（jdoc 渲染），编辑模式用 .oe-h-num（编辑器只读装饰）
        const numEl = el.querySelector('.docx-h-num, .oe-h-num');
        items.push({
            id,
            text: (el.textContent || '').replace(/\s+/g, ' ').trim(),
            level,
            number: numEl ? (numEl.textContent || '').trim() : ''
        });
    });

    return items;
}

export interface OutlineHandlers {
    /** 点击目录项：跳转并可选选中文本（编辑模式会顺带选中块） */
    onJump(item: OutlineItem, el: HTMLElement): void;
    /** 折叠 / 展开侧栏 */
    onToggle?(open: boolean): void;
}

export interface OutlineView {
    /** 侧栏根元素（挂到外壳左侧） */
    readonly el: HTMLElement;
    /** 用新的大纲重建条目 */
    update(items: OutlineItem[]): void;
    /** 高亮当前所在标题（滚动联动） */
    setActive(id: string | null): void;
    /** 当前高亮的条目 id */
    activeId(): string | null;
    destroy(): void;
}

/** 创建左侧目录栏 */
export function createOutline(handlers: OutlineHandlers): OutlineView {
    const list = document.createElement('nav');
    list.className = 'oe-outline-list';

    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'oe-outline-toggle';
    toggle.title = '折叠 / 展开目录';
    toggle.setAttribute('aria-label', '折叠或展开目录');

    const title = document.createElement('span');
    title.className = 'oe-outline-title';
    title.textContent = '目录';

    const head = document.createElement('div');
    head.className = 'oe-outline-head';
    head.append(toggle, title);

    const body = document.createElement('div');
    body.className = 'oe-outline-body';
    body.appendChild(list);

    const el = document.createElement('aside');
    el.className = 'oe-outline';
    el.append(head, body);

    let entries: { item: OutlineItem; node: HTMLElement }[] = [];
    let current: string | null = null;
    let collapsed = false;

    toggle.addEventListener('click', () => {
        collapsed = !collapsed;
        el.classList.toggle('collapsed', collapsed);
        handlers.onToggle?.(!collapsed);
    });

    const setActive = (id: string | null): void => {
        current = id;
        for (const entry of entries) entry.node.classList.toggle('active', entry.item.id === id);
        const hit = entries.find((entry) => entry.item.id === id);
        hit?.node.scrollIntoView({ block: 'nearest' });
    };

    return {
        el,
        update(items: OutlineItem[]): void {
            entries = [];
            list.textContent = '';
            for (const item of items) {
                const node = document.createElement('button');
                node.type = 'button';
                node.className = 'oe-outline-item';
                node.dataset.level = String(item.level);
                node.dataset.id = item.id;
                node.title = item.text;
                const label = document.createElement('span');
                label.className = 'oe-outline-text';
                // 有章节号时单独成段，便于等宽对齐编号列
                if (item.number) {
                    const num = document.createElement('span');
                    num.className = 'oe-outline-num';
                    num.textContent = item.number;
                    label.append(num);
                }
                label.append(document.createTextNode(item.text.replace(/^\s*\d+(\.\d+)*\.?\s*/, '')));
                node.appendChild(label);
                node.addEventListener('click', () => {
                    const target = document.querySelector<HTMLElement>(`[${ANCHOR_ATTR}="${item.id}"]`);
                    if (target) handlers.onJump(item, target);
                    setActive(item.id);
                });
                list.appendChild(node);
                entries.push({ item, node });
            }
            body.hidden = !items.length;
            el.classList.toggle('empty', !items.length);
            if (current) setActive(current);
        },
        setActive,
        activeId(): string | null {
            return current;
        },
        destroy(): void {
            list.textContent = '';
            entries = [];
        }
    };
}

/** 滚动联动的句柄 */
export interface ActiveHeadingTracker {
    /** 立即重算当前高亮（缩放 / 布局变化后调用，这些场景不触发 scroll） */
    refresh(): void;
    /** 解除监听 */
    dispose(): void;
}

/**
 * 滚动联动：滚动内容区时高亮当前视口顶部的标题
 *
 * @param scrollEl 内容滚动容器
 * @param ids 锚点 id 列表（与正文顺序一致）
 * @returns 句柄，可主动 refresh 或解除监听
 */
export function trackActiveHeading(
    scrollEl: HTMLElement,
    ids: string[],
    onChange: (id: string | null) => void
): ActiveHeadingTracker {
    if (!ids.length) return { refresh: () => undefined, dispose: () => undefined };
    let ticking = false;
    let last: string | null = null;

    const pick = (): void => {
        ticking = false;
        // 阈值用视觉像素：scrollEl 与标题元素的 rect 同处视觉坐标系，
        // 缩放后两者同步缩放，比较依然成立，无需再除 zoom。
        const top = scrollEl.getBoundingClientRect().top;
        let active: string | null = ids[0];
        for (const id of ids) {
            const el = scrollEl.querySelector<HTMLElement>(`[${ANCHOR_ATTR}="${id}"]`);
            if (!el) continue;
            if (el.getBoundingClientRect().top - top <= 24) active = id;
            else break;
        }
        if (active !== last) {
            last = active;
            onChange(active);
        }
    };

    const onScroll = (): void => {
        if (ticking) return;
        ticking = true;
        requestAnimationFrame(pick);
    };

    scrollEl.addEventListener('scroll', onScroll, { passive: true });
    return {
        refresh: () => pick(),
        dispose: () => scrollEl.removeEventListener('scroll', onScroll)
    };
}
