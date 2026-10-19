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
import { icon } from '../icons';
import type { DocFormat, EditorMode, HostToWebview, WebviewToHost } from '../../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let latest: Uint8Array = new Uint8Array(0);
let original: Uint8Array = new Uint8Array(0);
let gen = 0; // init 代数：丢弃过期的异步 applyMode 结果，防快速切换竞态

let editor: DocxEditor | XlsxEditor | null = null;
let root: HTMLElement;
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
            post({ type: 'sync', bytes: null, dirty: false });
        } else {
            latest = bytes;
            post({ type: 'sync', bytes: bytesToB64(bytes), dirty: true });
        }
    },
    onToast: (message: string, kind?: 'info' | 'warn' | 'error'): void => toast(message, kind)
};

// ---------- 外壳 ----------
function buildShell(): void {
    root = document.getElementById('app')!;
    root.innerHTML = '';
    const app = h('div', { class: 'oe' });

    // 不画顶栏：文档标题 / 未保存点由 VS Code 编辑器标题栏呈现；查看/编辑切换放在
    // VS Code 标题栏（editor/title 菜单），由 host 通过 setMode 消息驱动。
    toolbarEl = h('div', { class: 'oe-toolbar' });
    contentEl = h('div', { class: 'oe-content' });
    statusEl = h('span', { class: 'grow' });

    app.append(toolbarEl, contentEl, h('div', { class: 'oe-statusbar' }, [statusEl]));
    root.append(app);
}

function updateModeHint(): void {
    if (statusEl && mode === 'view') statusEl.textContent = '查看模式（只读）';
}

// ---------- 模式 ----------
async function applyMode(): Promise<void> {
    updateModeHint();
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
    } catch (err: unknown) {
        const msg = (err as { message?: string })?.message || String(err);
        contentEl.append(h('pre', { class: 'oe-error', text: '加载失败：' + msg }));
        post({ type: 'error', message: msg });
    }
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
        case 'setMode':
            // 来自 VS Code 标题栏的查看/编辑按钮
            if (msg.mode !== mode) {
                mode = msg.mode;
                post({ type: 'setMode', mode });
                await applyMode();
            }
            break;
        case 'bytes': {
            latest = b64ToBytes(msg.bytes);
            original = latest;
            if (mode === 'view') await enterView();
            break;
        }
        case 'saved':
            toast('已保存');
            break;
        case 'reverted':
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
    post({ type: 'ready' });
}

main();
