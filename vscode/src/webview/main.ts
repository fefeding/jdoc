import { pptxToHtml, docxToHtml, xlsxToHtml, pdfToJson } from '@fefeding/jdoc';
import { chartRenderer } from './chart-renderer';
import { b64ToBytes, escapeHtml } from './util';
import { baseCss } from './style';
import type { DocFormat, EditorMode, HostToWebview } from '../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let currentBytes: Uint8Array = new Uint8Array(0);

function applyBaseStyle(): void {
    const style = document.createElement('style');
    style.textContent = baseCss;
    document.head.appendChild(style);
}

function injectGlobalStyles(css: string): void {
    let el = document.getElementById('pptx-styles') as HTMLStyleElement | null;
    if (!el) {
        el = document.createElement('style');
        el.id = 'pptx-styles';
        document.head.appendChild(el);
    }
    el.textContent = css;
}

function buildToolbar(toolbar: HTMLElement): void {
    toolbar.innerHTML = '';
    const viewBtn = document.createElement('button');
    viewBtn.className = 'tbtn' + (mode === 'view' ? ' active' : '');
    viewBtn.textContent = '查看';
    viewBtn.onclick = () => setMode('view');
    toolbar.appendChild(viewBtn);

    const editBtn = document.createElement('button');
    editBtn.className = 'tbtn' + (mode === 'edit' ? ' active' : '');
    editBtn.textContent = '编辑';
    editBtn.onclick = () => setMode('edit');
    toolbar.appendChild(editBtn);
}

function renderShell(): { toolbar: HTMLElement; content: HTMLElement } {
    const app = document.getElementById('app')!;
    app.innerHTML = '';
    const toolbar = document.createElement('div');
    toolbar.className = 'toolbar';
    app.appendChild(toolbar);
    const content = document.createElement('div');
    content.id = 'content';
    app.appendChild(content);
    return { toolbar, content };
}

function setMode(next: EditorMode): void {
    mode = next;
    vscode.postMessage({ type: 'setMode', mode });
    const { toolbar } = renderShell();
    buildToolbar(toolbar);
    void renderContent();
}

async function renderContent(): Promise<void> {
    const content = document.getElementById('content')!;
    if (mode === 'edit') {
        // PPTX / DOCX / XLSX 已各自使用独立编辑器入口（pptx-editor.js / office-editor.js）；
        // 能走到这里的仅 PDF，按只读处理。
        content.innerHTML = '<div class="hint">PDF 为只读查看，暂不支持编辑。</div>';
        return;
    }
    try {
        if (format === 'pptx') await renderPptx(content, currentBytes);
        else if (format === 'docx') await renderDocx(content, currentBytes);
        else if (format === 'xlsx') await renderXlsx(content, currentBytes);
        else if (format === 'pdf') await renderPdf(content, currentBytes);
    } catch (err: any) {
        content.innerHTML = `<pre class="error">渲染失败：${escapeHtml(String(err?.message || err))}</pre>`;
        vscode.postMessage({ type: 'error', message: String(err?.message || err) });
    }
}

async function renderPptx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const r = await pptxToHtml(bytes, { slideMode: false, keyBoardShortCut: false });
    injectGlobalStyles(r.styles?.global || '');
    const wrap = document.createElement('div');
    wrap.className = 'pptx-slides';
    for (const slide of r.slides || []) {
        const holder = document.createElement('div');
        holder.innerHTML = slide.html;
        const scaler = document.createElement('div');
        scaler.className = 'slide-scaler';
        while (holder.firstChild) scaler.appendChild(holder.firstChild);
        wrap.appendChild(scaler);
    }
    content.appendChild(wrap);
    // 图表：ECharts 已由 media 脚本注入为全局变量，chart-renderer 直接使用
    if (r.charts && r.charts.length) {
        chartRenderer.renderCharts(r.charts, wrap);
    }
}

async function renderDocx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await docxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'docx-body';
    body.innerHTML = html;
    content.appendChild(body);
}

async function renderXlsx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await xlsxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'xlsx-body';
    body.innerHTML = html;
    content.appendChild(body);
}

async function renderPdf(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    // webview 是单一打包产物，jdoc 内部用 import.meta.url 派生的 pdfjs worker 地址无效；
    // 这里提前设置有效的 worker URL（extension 注入到 window.__JDC_PDF_WORKER__）。
    // 利用 ESM 模块缓存：本处与 jdoc 内部 import 同一 pdfjs 实例，workerSrc 设置后即生效。
    if (typeof window !== 'undefined' && (window as any).__JDC_PDF_WORKER__) {
        try {
            const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
            if (pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
                pdfjs.GlobalWorkerOptions.workerSrc = (window as any).__JDC_PDF_WORKER__;
            }
        } catch {
            // 设置失败不阻断：jdoc 会回退到自身逻辑并给出可读错误
        }
    }
    const result = await pdfToJson(bytes);
    const pages = (result.pages || []) as Array<{ width?: number; height?: number; text?: string }>;
    for (const page of pages) {
        const p = document.createElement('div');
        p.className = 'pdf-page';
        const w = page.width || 595;
        const h = page.height || 842;
        p.style.width = `${w}px`;
        p.style.minHeight = `${h}px`;
        const text = document.createElement('div');
        text.className = 'page-text';
        text.textContent = page.text || '';
        p.appendChild(text);
        content.appendChild(p);
    }
    if (!pages.length) content.innerHTML = '<div class="hint">无页面内容</div>';
}

window.addEventListener('message', (e) => {
    const msg = e.data as HostToWebview;
    switch (msg.type) {
        case 'init':
            format = msg.format;
            mode = msg.mode || 'view';
            currentBytes = b64ToBytes(msg.bytes);
            applyBaseStyle();
            {
                const { toolbar } = renderShell();
                buildToolbar(toolbar);
            }
            void renderContent();
            vscode.postMessage({ type: 'ready' });
            break;
        case 'reverted':
            void renderContent();
            break;
        case 'saved':
        case 'info':
            break;
    }
});
