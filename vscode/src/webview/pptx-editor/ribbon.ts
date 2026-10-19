// PowerPoint-style ribbon: title bar + quick access, tab strip, and per-tab control groups.
//
// The ribbon owns no state — `main.ts` computes a `RibbonState` on every render and calls
// `sync()`. Each control declares:
//   • `needs`  — capability flags (from `state.caps`) that must be true for it to be enabled
//   • `active` — predicate for the "pressed / selected" look (bold, align, grid, …)
//   • `onSync` — pushes a current value into a select / colour swatch
// Everything else (labels, tooltips, dropdown contents) resolves through `i18n` here, so
// switching language is just a matter of rebuilding the ribbon (see `main.ts`).
//
// The controls of each tab live in `ribbon-tabs.ts`; this file only provides the builders.
import { h } from './util';
import { icon } from './icons';
import { t, tip, withAccel } from './i18n';
import { openMenuAt, MenuItem } from './menu';
import { THEMES } from '@fefeding/jdoc';
import type { EditorMode } from './protocol';
import { buildTabs, RibbonTab } from './ribbon-tabs';

export type RibbonState = {
  mode: EditorMode;
  canUndo: boolean;
  canRedo: boolean;
  zoom: number;
  grid: boolean;
  snap: boolean;
  presenting: boolean;
  slideCount: number;
  slideIndex: number;
  slideHidden: boolean;
  theme: string;
  sizeKey: string;
  transition: { type: string; duration: number; advanceOnClick: boolean };
  animations: { index: number; text: string }[];
  /** Named capability flags consumed by `needs` (see `ribbonCaps` in main.ts). */
  caps: Record<string, boolean | string>;
};

export type RibbonApi = {
  /** Invoke a `main.ts` command by name; returns whatever the command returns (menus). */
  cmd(name: string, ...args: any[]): any;
};

type Widget = { el: HTMLElement; sync?: (s: RibbonState) => void };

const COLLAPSE_KEY = 'pptx.ribbon.collapsed';
const ACTIVE_TAB_KEY = 'pptx.ribbon.tab';

/** Standard colours offered by the colour dropdowns (Office's first row). */
export const PALETTE = [
  '#ffffff', '#000000', '#eeece1', '#1f497d', '#4f81bd', '#c0504d',
  '#9bbb59', '#8064a2', '#4bacc6', '#f79646', '#e5e5e5', '#7f7f7f'
];

export function themeAccents(themeId: string): string[] {
  const th: any = THEMES.find((x: any) => x.id === themeId) || THEMES[0];
  return [th.accent, ...(th.accents || [])].filter(Boolean);
}

export function readStore(key: string): string | null {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}
export function writeStore(key: string, value: string): void {
  try {
    localStorage.setItem(key, value);
  } catch {
    /* storage disabled (private mode) — the ribbon still works, it just won't remember */
  }
}

/** Label + platform shortcut hint, used for every tooltip in the ribbon. */
export { withAccel };

export function createRibbon(api: RibbonApi): { root: HTMLElement; sync(s: RibbonState): void } {
  const widgets: Widget[] = [];
  let lastState: RibbonState | null = null;

  const cap = (s: RibbonState, needs?: string | string[]): boolean => {
    if (!needs) return true;
    const list = Array.isArray(needs) ? needs : [needs];
    return list.every((n) => !!s.caps[n]);
  };
  const applyDisabled = (el: HTMLElement, s: RibbonState, needs?: string | string[]) => {
    const on = cap(s, needs);
    el.classList.toggle('disabled', !on);
    el.setAttribute('aria-disabled', on ? 'false' : 'true');
  };

  // ---------- control builders ----------
  /** Big button (icon above a label) — a group's primary action, as in Office. */
  function bigButton(o: {
    ico: string;
    label: string;
    title?: string;
    accel?: string;
    needs?: string | string[];
    onClick: () => void;
  }): Widget {
    const el = h('button', { class: 'rb-btn rb-btn-big', type: 'button' }, [icon(o.ico, 20), h('span', {}, [o.label])]);
    el.title = o.title || withAccel(o.label, o.accel);
    el.setAttribute('aria-label', o.label);
    el.onclick = () => {
      if (!el.classList.contains('disabled')) o.onClick();
    };
    widgets.push({ el, sync: (s) => applyDisabled(el, s, o.needs) });
    return { el };
  }

  /** Small icon-only button (align / arrange / toggle controls). */
  function miniButton(o: {
    ico?: string;
    glyph?: string;
    glyphStyle?: string;
    title: string;
    needs?: string | string[];
    active?: (s: RibbonState) => boolean;
    onClick: () => void;
  }): Widget {
    const el = h('button', { class: 'rb-btn rb-btn-mini', type: 'button' });
    if (o.glyph) el.append(h('span', { class: 'rb-glyph', style: o.glyphStyle || '' }, [o.glyph]));
    else el.append(icon(o.ico || 'shape', 16));
    el.title = o.title;
    el.setAttribute('aria-label', o.title);
    el.onclick = () => {
      if (!el.classList.contains('disabled')) o.onClick();
    };
    widgets.push({ el, sync: (s) => { applyDisabled(el, s, o.needs); el.classList.toggle('on', !!o.active?.(s)); } });
    return { el };
  }

  /** Big button + dropdown arrow — Office's split button (New Slide ▾, Layout ▾ …). */
  function splitButton(o: {
    ico: string;
    label: string;
    accel?: string;
    needs?: string | string[];
    onClick: () => void;
    menu: () => MenuItem[];
  }): Widget {
    const wrap = h('div', { class: 'rb-split' });
    const main = h('button', { class: 'rb-btn rb-btn-big rb-split-main', type: 'button' }, [icon(o.ico, 20), h('span', {}, [o.label])]);
    const arrow = h('button', { class: 'rb-btn rb-split-arrow', type: 'button', 'aria-haspopup': 'menu' }, [icon('chevron', 12)]);
    main.title = withAccel(o.label, o.accel);
    arrow.title = o.label;
    main.onclick = () => {
      if (!main.classList.contains('disabled')) o.onClick();
    };
    arrow.onclick = () => {
      if (!arrow.classList.contains('disabled')) openMenuAt(arrow, o.menu(), { minWidth: 200 });
    };
    wrap.append(main, arrow);
    widgets.push({ el: main, sync: (s) => applyDisabled(main, s, o.needs) });
    widgets.push({ el: arrow, sync: (s) => applyDisabled(arrow, s, o.needs) });
    return { el: wrap };
  }

  /** Caption + select (font family, size, spacing, transition speed, …). */
  function selectBox(o: {
    label: string;
    options: { value: string; label: string }[];
    width?: number;
    needs?: string | string[];
    onChange: (v: string) => void;
    onSync?: (s: RibbonState) => string;
  }): Widget & { setOptions(options: { value: string; label: string }[]): void } {
    const sel = h('select', { class: 'rb-select' });
    if (o.width) sel.style.width = o.width + 'px';
    const setOptions = (options: { value: string; label: string }[]) => {
      sel.innerHTML = '';
      for (const opt of options) sel.append(h('option', { value: opt.value }, [opt.label]));
    };
    setOptions(o.options);
    sel.title = o.label;
    sel.addEventListener('change', () => o.onChange(sel.value));
    const el = h('div', { class: 'rb-field' }, [h('label', {}, [o.label]), sel]);
    widgets.push({
      el,
      sync: (s) => {
        applyDisabled(sel, s, o.needs);
        const v = o.onSync?.(s);
        // Don't fight the user while the dropdown is open, and keep the first option (rather
        // than blanking it) when there is no value to show — e.g. nothing selected yet.
        if (v && document.activeElement !== sel && sel.value !== v) sel.value = v;
      }
    });
    return { el, setOptions };
  }

  /** Swatch + dropdown — Office-style colour picker. */
  function colorBox(o: {
    title: string;
    ico: string;
    needs?: string | string[];
    onInput: (v: string) => void;
    onSync?: (s: RibbonState) => string;
  }): Widget {
    const input = h('input', { type: 'color', class: 'rb-color-input' });
    input.addEventListener('input', () => o.onInput(input.value));
    const swatch = h('span', { class: 'rb-color-swatch' });
    const btn = h('button', { class: 'rb-btn rb-btn-color', type: 'button', title: o.title, 'aria-label': o.title }, [
      h('span', { class: 'rb-color-stack' }, [icon(o.ico, 16), swatch]),
      h('span', { class: 'rb-caret' }, [icon('chevron', 10)])
    ]);
    btn.onclick = () => {
      const current = (lastState && o.onSync?.(lastState)) || '#ffffff';
      const items: MenuItem[] = [
        { type: 'label', label: o.title },
        { type: 'item', label: t('common.none'), icon: 'close', onClick: () => o.onInput('none') }
      ];
      const swatches = (colors: string[], caption: string) => {
        items.push({ type: 'label', label: caption });
        for (const c of colors) {
          items.push({
            type: 'item',
            label: c.toUpperCase(),
            checked: c.toLowerCase() === current.toLowerCase(),
            onClick: () => o.onInput(c)
          });
        }
      };
      if (lastState) swatches(themeAccents(lastState.theme), t('cmd.theme'));
      swatches(PALETTE, t('insp.color'));
      openMenuAt(btn, items, { minWidth: 152 });
    };
    widgets.push({
      el: btn,
      sync: (s) => {
        applyDisabled(btn, s, o.needs);
        const v = o.onSync?.(s);
        if (!v) return;
        const hex = /^#[0-9a-f]{6}$/i.test(v) ? v : '#ffffff';
        if (input.value !== hex) input.value = hex;
        swatch.style.background = hex;
      }
    });
    return { el: btn };
  }

  /** Single-line horizontal gallery (transitions, animation presets, themes). */
  function gallery<T>(
    items: T[],
    renderItem: (item: T) => { el: HTMLElement; isActive?: (s: RibbonState) => boolean },
    opts: { needs?: string | string[]; caption?: string } = {}
  ): Widget {
    const list = h('div', { class: 'rb-gallery' });
    const entries: { el: HTMLElement; isActive?: (s: RibbonState) => boolean }[] = [];
    for (const item of items) {
      const r = renderItem(item);
      list.append(r.el);
      entries.push(r);
    }
    widgets.push({
      el: list,
      sync: (s) => {
        applyDisabled(list, s, opts.needs);
        for (const e of entries) e.el.classList.toggle('on', !!e.isActive?.(s));
      }
    });
    const el = h('div', { class: 'rb-stack' }, opts.caption ? [h('span', { class: 'rb-caption' }, [opts.caption]), list] : [list]);
    return { el };
  }

  function chip(
    label: string,
    title: string,
    onClick: () => void,
    active?: (s: RibbonState) => boolean
  ): { el: HTMLElement; isActive?: (s: RibbonState) => boolean } {
    const el = h('button', { class: 'rb-chip', type: 'button', title }, [label]);
    el.onclick = onClick;
    return { el, isActive: active };
  }

  /** A titled group of control rows (Office ribbon group box). */
  function group(title: string, rows: (Widget | null)[][], editOnly = false): HTMLElement {
    const body = h('div', { class: 'rb-group-body' });
    for (const row of rows) {
      if (!row.length) continue;
      const r = h('div', { class: 'rb-row' });
      for (const item of row) if (item) r.append(item.el);
      body.append(r);
    }
    return h('div', { class: 'rb-group' + (editOnly ? ' edit-only' : '') }, [body, h('div', { class: 'rb-group-title' }, [title])]);
  }

  function advanceCheck(): Widget {
    const box = h('input', { type: 'checkbox' });
    box.addEventListener('change', () => api.cmd('setTransAdvance', box.checked));
    const el = h('label', { class: 'rb-check', title: t('insp.clickToAdvance') }, [box, h('span', {}, [t('insp.clickToAdvance')])]);
    widgets.push({
      el,
      sync: (s) => {
        applyDisabled(box, s, 'edit');
        if (document.activeElement !== box) box.checked = s.transition.advanceOnClick !== false;
      }
    });
    return { el };
  }

  function animListBox(): Widget {
    const el = h('div', { class: 'rb-anim-list' });
    widgets.push({
      el,
      sync: (s) => {
        el.innerHTML = '';
        applyDisabled(el, s, 'edit');
        if (!s.animations.length) {
          el.append(h('div', { class: 'empty' }, [t('insp.noAnimations')]));
          return;
        }
        for (const a of s.animations) {
          el.append(
            h('div', { class: 'rb-anim-row' }, [
              h('span', { class: 'rb-anim-text', title: a.text }, [a.text]),
              h('button', { class: 'rb-btn rb-btn-mini', title: t('insp.animMoveUp'), onclick: () => api.cmd('animMove', a.index, -1) }, [icon('up', 12)]),
              h('button', { class: 'rb-btn rb-btn-mini', title: t('insp.animMoveDown'), onclick: () => api.cmd('animMove', a.index, 1) }, [icon('down', 12)]),
              h('button', { class: 'rb-btn rb-btn-mini', title: t('insp.animRemove'), onclick: () => api.cmd('animRemove', a.index) }, [icon('close', 12)])
            ])
          );
        }
      }
    });
    return { el };
  }

  const builders = { bigButton, miniButton, splitButton, selectBox, colorBox, gallery, chip, group, advanceCheck, animListBox };

  // ---------- title bar / quick access ----------
  // 文档标题 / 未保存点 / 查看-编辑切换都不在这里画：前两者与 VS Code 编辑器标题栏重复，
  // 模式切换与文档切换由 host 注册为 editor/title 命令（见 package.json 的 menus）。
  const collapseBtn = h('button', { class: 'rb-btn rb-btn-mini rb-collapse', type: 'button', 'aria-expanded': 'true' }, [icon('chevron', 12)]);
  collapseBtn.title = t('tbar.ribbonHide');

  // Preview mode has few actions, so it shows a compact shortcut menu in the header instead of the
  // expandable tab ribbon. (Edit mode keeps the full ribbon.)
  const previewMenuBtn = h('button', { class: 'rb-btn rb-btn-mini rb-preview-menu', type: 'button', title: t('tbar.menu') }, [icon('menu', 16)]);
  previewMenuBtn.onclick = () => {
    const items: MenuItem[] = [
      { type: 'item', label: t('cmd.present'), icon: 'present', accel: 'F5', onClick: () => api.cmd('present') },
      { type: 'item', label: t('cmd.presentHere'), icon: 'presentHere', accel: 'Shift+F5', onClick: () => api.cmd('presentHere') },
      { type: 'sep' },
      { type: 'item', label: t('cmd.fit'), icon: 'fit', onClick: () => api.cmd('zoom', 'fit') },
      { type: 'item', label: t('cmd.zoomActual'), icon: 'slideSize', onClick: () => api.cmd('zoom', 'actual') },
      { type: 'sep' },
      { type: 'item', label: t('cmd.shortcuts'), icon: 'help', onClick: () => api.cmd('shortcuts') }
    ];
    openMenuAt(previewMenuBtn, items, { minWidth: 210 });
  };

  // Tabs render inside the title bar so the ribbon costs only one extra row instead of two.
  const tabsEl = h('div', { class: 'rb-tabs', role: 'tablist' });
  const bodyEl = h('div', { class: 'rb-body' });
  const tabs: RibbonTab[] = buildTabs(builders, api);
  const tabBtns = new Map<string, HTMLElement>();

  for (const tab of tabs) {
    const btn = h('button', { class: 'rb-tab', type: 'button', role: 'tab', 'data-tab': tab.id }, [tab.label]);
    btn.onclick = () => {
      expandForTabClick();
      setTab(tab.id);
    };
    tabsEl.append(btn);
    tabBtns.set(tab.id, btn);
    bodyEl.append(tab.panel);
  }

  function isVisible(tab: RibbonTab, edit: boolean): boolean {
    return !tab.editOnly || edit;
  }
  // The active tab is remembered per mode: preview only shows the show / view / help tabs,
  // so switching back to edit mode should not keep the user on "Slide Show".
  const tabKey = (mode: EditorMode) => `${ACTIVE_TAB_KEY}.${mode}`;
  /** `persist` is false for the sync-time repair pass, which runs on every render. */
  function setTab(id: string, persist = true): void {
    for (const tab of tabs) {
      const on = tab.id === id;
      tab.panel.style.display = on ? '' : 'none';
      tabBtns.get(tab.id)?.classList.toggle('active', on);
    }
    if (persist) writeStore(tabKey(lastState?.mode || 'preview'), id);
  }
  /** Picking a tab by hand while the ribbon is collapsed should reveal its command group. */
  function expandForTabClick(): void {
    if (!root.classList.contains('rb-collapsed')) return;
    root.classList.remove('rb-collapsed');
    collapseBtn.classList.remove('flip');
    collapseBtn.setAttribute('aria-expanded', 'true');
    collapseBtn.title = t('tbar.ribbonHide');
    writeStore(COLLAPSE_KEY, '0');
  }

  const rightQat = h('div', { class: 'rb-qat' }, [
    miniButton({ ico: 'help', title: withAccel(t('cmd.shortcuts'), '?'), onClick: () => api.cmd('shortcuts') }).el
  ]);

  const titlebar = h('div', { class: 'rb-titlebar' }, [
    h('div', { class: 'rb-qat' }, [
      miniButton({ ico: 'undo', title: withAccel(t('cmd.undo'), 'mod+z'), needs: 'canUndo', onClick: () => api.cmd('undo') }).el,
      miniButton({ ico: 'redo', title: withAccel(t('cmd.redo'), 'mod+y'), needs: 'canRedo', onClick: () => api.cmd('redo') }).el,
      miniButton({ ico: 'present', title: withAccel(t('tbar.present'), 'f5'), onClick: () => api.cmd('present') }).el
    ]),
    previewMenuBtn,
    tabsEl,
    h('div', { class: 'rb-spring' }),
    rightQat,
    collapseBtn
  ]);

  const root = h('div', { class: 'ribbon' }, [titlebar, bodyEl]);

  // The first sync() picks the right tab for the current mode; start on the first one.
  setTab(tabs[0].id, false);

  collapseBtn.onclick = () => {
    const collapsed = root.classList.toggle('rb-collapsed');
    collapseBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    collapseBtn.classList.toggle('flip', collapsed);
    collapseBtn.title = collapsed ? t('tbar.ribbonShow') : t('tbar.ribbonHide');
    writeStore(COLLAPSE_KEY, collapsed ? '1' : '0');
  };
  // The ribbon starts collapsed (one-line header only); the user expands it by clicking a tab or
  // the toggle. An explicit "expand" preference (COLLAPSE_KEY === '0') is remembered.
  if (readStore(COLLAPSE_KEY) !== '0') {
    root.classList.add('rb-collapsed');
    collapseBtn.classList.add('flip');
    collapseBtn.setAttribute('aria-expanded', 'false');
    collapseBtn.title = t('tbar.ribbonShow');
  }

  function sync(s: RibbonState): void {
    const modeChanged = lastState?.mode !== s.mode;
    lastState = s;
    // Preview mode swaps the tab ribbon for a compact header menu (no expandable toolbar).
    root.classList.toggle('rb-preview', s.mode === 'preview');
    for (const w of widgets) w.sync?.(s);
    // Preview mode has no editing tabs — hide the ones that would be empty and fall back to
    // the first visible tab. Visibility is re-applied on every render because switching back
    // to edit mode must restore the panels that were hidden while previewing.
    const edit = s.mode === 'edit';
    const visible = tabs.filter((tab) => isVisible(tab, edit));
    for (const tab of tabs) {
      tabBtns.get(tab.id)!.style.display = isVisible(tab, edit) ? '' : 'none';
    }
    const activeId = modeChanged
      ? readStore(tabKey(s.mode))
      : tabs.find((tab) => tabBtns.get(tab.id)!.classList.contains('active'))?.id;
    const target = visible.find((tab) => tab.id === activeId) || visible[0];
    if (target) setTab(target.id);
  }

  return { root, sync };
}