// office-editor 入口：DOCX / XLSX 的查看（view）与编辑（edit）。
// 独立自包含入口，避免与查看版 webview.js 的双 message 监听 / 双 acquireVsCodeApi 冲突。
//
// 数据流：host 发 init(bytes) → view 用 docxToHtml/xlsxToHtml 只读渲染；edit 用块式/网格
// 编辑器载入 blocks/sheets 模型 → 编辑 → jsonToDocx/jsonToXlsx 生成字节 → sync 回 host 保存。
import { docxToHtml, xlsxToHtml } from '@fefeding/jdoc';
import { officeCss } from './style';
import { CHROME_CSS } from '../chrome';
import { toast } from './framework';
import { createDocxEditor, type DocxEditor } from './docx';
import { createXlsxEditor, type XlsxEditor } from './xlsx';
import { b64ToBytes, bytesToB64, h, escapeHtml } from './util';
import { icon, iconNode } from '../icons';
import { createTopBar, createDocTabs, type TopBar, type DocTabs } from '../chrome';
import type { DocFormat, EditorMode, HostToWebview, WebviewToHost } from '../../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let docTitle = 'document';
let dirty = false;
let latest: Uint8Array = new Uint8Array(0);
let original: Uint8Array = new Uint8Array(0);
let gen = 0; // init 代数：丢弃过期的异步 applyMode 结果，防快速切换竞态

let editor: DocxEditor | XlsxEditor | null = null;
let root: HTMLElement;
let topBar: TopBar | null = null;
let docTabs: DocTabs | null = null;
let titleEl: HTMLElement;
let dirtyEl: HTMLElement;
let toolbarEl: HTMLElement;
let contentEl: HTMLElement;
let statusEl: HTMLElement;

function post(msg: WebviewToHost): void {
    vscode.postMessage(msg);
}

const hooks = {
    onSync: (bytes: Uint8Array | null): void => {
        if (bytes === null) {
            // 放弃编辑回到导入态
            latest = original;
            dirty = false;
            post({ type: 'sync', bytes: null, dirty: false });
        } else {
            latest = bytes;
            dirty = true;
            post({ type: 'sync', bytes: bytesToB64(bytes), dirty: true });
        }
        updateTitle();
    },
    onToast: (message: string, kind?: 'info' | 'warn' | 'error'): void => toast(message, kind)
};

// ---------- 外壳 ----------
function buildShell(): void {
    root = document.getElementById('app')!;
    root.innerHTML = '';
    const app = h('div', { class: 'oe' });

    topBar = createTopBar(
        [
            { id: 'view', label: '查看', icon: 'eye' },
            { id: 'edit', label: '编辑', icon: 'pencil' }
        ],
        mode,
        (id) => setMode(id as EditorMode),
        { node: iconNode }
    );
    topBar.setBadge(format === 'xlsx' ? 'XLSX' : 'DOCX');
    titleEl = topBar.nameEl;
    dirtyEl = topBar.dirtyEl;
    // 右上角文档切换标签（由 host 的 docs 广播填充）
    docTabs = createDocTabs((uri) => post({ type: 'openDocument', uri }));
    topBar.el.append(docTabs.el);
    const top = topBar.el;

    toolbarEl = h('div', { class: 'oe-toolbar' });
    contentEl = h('div', { class: 'oe-content' });
    statusEl = h('span', { class: 'grow' });

    app.append(top, toolbarEl, contentEl, h('div', { class: 'oe-statusbar' }, [statusEl]));
    root.append(app);
}

function updateTitle(): void {
    topBar?.setTitle(docTitle, dirty);
}

function updateSeg(): void {
    topBar?.setActive(mode);
    if (statusEl && mode === 'view') statusEl.textContent = '查看模式（只读）';
}

// ---------- 模式 ----------
async function applyMode(): Promise<void> {
    updateSeg();
    updateTitle();
    if (mode === 'edit') await enterEdit();
    else await enterView();
}

async function enterView(): Promise<void> {
    // 离开编辑：先把编辑产出同步为字节，再忠实渲染
    if (editor) {
        await editor.syncNow();
        editor.destroy();
        editor = null;
    }
    toolbarEl.innerHTML = '';
    statusEl.textContent = '查看模式（只读）';
    contentEl.innerHTML = '';
    try {
        let html = '';
        if (format === 'docx') html = (await docxToHtml(latest)).html;
        else if (format === 'xlsx') html = (await xlsxToHtml(latest)).html;
        else html = `<div class="oe-hint">office-editor 仅支持 DOCX / XLSX（当前：${escapeHtml(format)}）</div>`;
        const view = h('div', { class: 'oe-view' }, [h('div', { class: 'page', html })]);
        contentEl.append(view);
    } catch (err: unknown) {
        const msg = (err as { message?: string })?.message || String(err);
        contentEl.append(h('pre', { class: 'oe-error', text: '渲染失败：' + msg }));
        post({ type: 'error', message: msg });
    }
}

async function enterEdit(): Promise<void> {
    contentEl.innerHTML = '';
    toolbarEl.innerHTML = '';
    if (editor) {
        editor.destroy();
        editor = null;
    }
    try {
        if (format === 'docx') {
            editor = createDocxEditor(hooks);
            statusEl.textContent = '';
        } else if (format === 'xlsx') {
            const x = createXlsxEditor(hooks);
            x.setStatusHandler((s) => (statusEl.textContent = s));
            editor = x;
        } else {
            contentEl.append(h('div', { class: 'oe-hint', text: `编辑仅支持 DOCX / XLSX（当前：${format}）` }));
            return;
        }
        editor.mount(toolbarEl, contentEl);
        await editor.load(latest);
        updateTitle();
    } catch (err: unknown) {
        const msg = (err as { message?: string })?.message || String(err);
        contentEl.append(h('pre', { class: 'oe-error', text: '加载失败：' + msg }));
        post({ type: 'error', message: msg });
    }
}

function setMode(next: EditorMode): void {
    if (next === mode) return;
    mode = next;
    post({ type: 'setMode', mode });
    void applyMode();
}

// ---------- 消息 ----------
window.addEventListener('message', async (event: MessageEvent<HostToWebview>) => {
    const msg = event.data;
    if (!msg || typeof msg !== 'object') return;
    switch (msg.type) {
        case 'init': {
            const my = ++gen; // 丢弃此后被更新的 init 所取代的异步结果
            format = msg.format;
            mode = msg.mode === 'edit' ? 'edit' : 'view';
            docTitle = msg.title || docTitle;
            dirty = msg.dirty;
            latest = b64ToBytes(msg.bytes);
            original = latest;
            if (editor) {
                editor.destroy();
                editor = null;
            }
            await applyMode();
            if (my !== gen) return;
            break;
        }
        case 'bytes': {
            latest = b64ToBytes(msg.bytes);
            original = latest;
            dirty = msg.dirty;
            if (mode === 'view') await enterView();
            updateTitle();
            break;
        }
        case 'docs':
            docTabs?.update(msg.docs, msg.activeUri);
            break;
        case 'saved':
            dirty = false;
            updateTitle();
            toast('已保存');
            break;
        case 'reverted':
            dirty = false;
            updateTitle();
            break;
        case 'info':
            toast(msg.message, msg.kind || 'info');
            break;
    }
});

// ---------- 启动 ----------
function main(): void {
    const style = document.createElement('style');
    style.textContent = CHROME_CSS + officeCss;
    document.head.appendChild(style);
    buildShell();
    updateSeg();
    post({ type: 'ready' });
}

main();
