import { pptxToHtml, docxToHtml, xlsxToHtml, markdownToHtml, textToHtml, renderMarkdownExtras } from '@fefeding/jdoc';
import { chartRenderer } from './chart-renderer';
import { b64ToBytes, bytesToB64, escapeHtml } from './util';
import { baseCss } from './style';
import { CHROME_CSS, createStatusBar } from './chrome';
import { createZoom, createZoomCluster } from './zoom';
import { createPdfViewer, PDF_CSS, type PdfViewer } from './pdf-viewer';
import type { DocFormat, EditorMode, HostToWebview } from '../protocol';

declare function acquireVsCodeApi(): { postMessage(msg: unknown): void };
const vscode = acquireVsCodeApi();

let format: DocFormat = 'pptx';
let mode: EditorMode = 'view';
let currentBytes: Uint8Array = new Uint8Array(0);
let statusText: HTMLElement | null = null;
/** 只握手一次：host 收到 ready 会再发一次 init，无条件回 ready 会形成无限循环。 */
let readySent = false;

/** 缩放：PDF / Markdown / 纯文本等只读内容共用，默认适应宽度，宽内容不会出屏外 */
const zoom = createZoom();

/** 把缩放绑定到当前渲染出的内容根 */
function bindZoom(): void {
    const content = document.getElementById('content');
    if (!content) return;
    const el = content.firstElementChild as HTMLElement | null;
    // 提示 / 错误块无需缩放
    if (!el || el.classList.contains('hint') || el.classList.contains('error')) {
        zoom.bind({ target: null, container: content });
        return;
    }
    zoom.bind({ target: el, container: content });
}

function applyBaseStyle(): void {
    const style = document.createElement('style');
    style.textContent = CHROME_CSS + baseCss + PDF_CSS;
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

/** 更新底部状态栏文案 */
function setStatus(text: string): void {
    if (statusText) statusText.textContent = text;
}

/** PDF 阅读器实例（仅 pdf 格式创建） */
let pdfViewer: PdfViewer | null = null;

function renderShell(): { content: HTMLElement } {
    const app = document.getElementById('app')!;
    // 重建外壳前先拆旧实例：PDF 编辑层在 document 上挂了 pointer 监听，
    // 不 close 会随每次 init 累积一份
    pdfViewer?.close();
    app.innerHTML = '';

    // PDF 有自己的阅读器外壳（顶部工具栏 + 左侧缩略图栏 + 正文），缩放也由它绑定
    if (format === 'pdf') {
        app.className = 'jc-app pdf-app';
        const status = createStatusBar();
        statusText = status.textEl;
        pdfViewer = createPdfViewer(app, zoom, setStatus, (editing) => {
            // 阅读器里手动开关编辑模式，同步给宿主，标题栏的「查看/编辑」命令才不会失配
            vscode.postMessage({ type: 'setMode', mode: editing ? 'edit' : 'view' });
        });
        // 叠加编辑产出新字节：先回传字节（标脏），再请求落盘，
        // 一步完成「生成 + 写回原文件」，不必让用户再按一次保存。
        pdfViewer.onSave((bytes) => {
            vscode.postMessage({ type: 'sync', bytes: bytesToB64(bytes), dirty: true });
            vscode.postMessage({ type: 'save' });
        });
        app.append(status.el);
        return { content: pdfViewer.content };
    }

    // 不画顶栏：标题 / 面包屑 / 未保存点由 VS Code 编辑器标题栏呈现，webview 再画一份只会重复。
    const content = document.createElement('div');
    content.className = 'jc-content';
    content.id = 'content';
    const status = createStatusBar();
    statusText = status.textEl;
    app.className = 'jc-app';
    app.append(content, status.el);
    // 缩放簇挂在状态栏尾部（createStatusBar 里已有一个 grow 占位，簇会自然靠右）
    status.el.append(createZoomCluster(zoom));
    return { content };
}

async function renderContent(): Promise<void> {
    const content = document.getElementById('content')!;
    // PDF 无论查看还是编辑都走阅读器外壳，编辑模式只是叠加一层 overlay 工具层
    if (format === 'pdf' && pdfViewer) {
        try {
            await pdfViewer.open(currentBytes);
            pdfViewer.setEditing(mode === 'edit', false);
        } catch (err: any) {
            pdfViewer.content.innerHTML = `<pre class="error">渲染失败：${escapeHtml(String(err?.message || err))}</pre>`;
            setStatus('渲染失败');
            vscode.postMessage({ type: 'error', message: String(err?.message || err) });
        }
        return;
    }
    if (mode === 'edit') {
        // PPTX / DOCX / XLSX 已各自使用独立编辑器入口（pptx-editor.js / office-editor.js）；
        // 能走到这里的仅 MD / TXT，按只读处理。
        const name = format === 'md' ? 'Markdown' : '纯文本';
        const hint = document.createElement('div');
        hint.className = 'hint';
        hint.textContent = `${name}为只读查看，暂不支持编辑。`;
        content.replaceChildren(hint);
        setStatus('只读预览');
        bindZoom();
        return;
    }
    try {
        if (format === 'pdf') {
            // 上面已处理；走到这里说明外壳没建成（理论上不会发生）
            return;
        }
        if (format === 'pptx') await renderPptx(content, currentBytes);
        else if (format === 'docx') await renderDocx(content, currentBytes);
        else if (format === 'xlsx') await renderXlsx(content, currentBytes);
        else if (format === 'md') await renderMarkdown(content, currentBytes);
        else if (format === 'txt') await renderText(content, currentBytes);
        bindZoom();
    } catch (err: any) {
        content.innerHTML = `<pre class="error">渲染失败：${escapeHtml(String(err?.message || err))}</pre>`;
        setStatus('渲染失败');
        bindZoom();
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
    const slideCount = (r.slides || []).length;
    setStatus(`共 ${slideCount} 页幻灯片${r.charts?.length ? `，${r.charts.length} 个图表` : ''}`);
}

async function renderDocx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await docxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'docx-body';
    body.innerHTML = html;
    content.appendChild(body);

    setStatus('DOCX 文档预览');
}

async function renderXlsx(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await xlsxToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'xlsx-body';
    body.innerHTML = html;
    content.appendChild(body);

    setStatus('XLSX 表格预览');
}

async function renderMarkdown(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html, messages: baseMessages } = await markdownToHtml(bytes);
    let messages = baseMessages || [];
    const body = document.createElement('div');
    body.className = 'md-body';
    body.innerHTML = html;
    content.appendChild(body);
    // 外部环境链接在新标签页打开（webview 内无法跳转）
    body.querySelectorAll<HTMLAnchorElement>('a[href^="http"]').forEach((a) => {
        a.setAttribute('target', '_blank');
        a.setAttribute('rel', 'noopener noreferrer');
    });
    // 增强内容：Mermaid 图表 / KaTeX 公式（动态加载依赖，失败时保留源码降级显示）
    try {
        // 主题判定：优先用 VS Code 给 webview body 加的类（实时跟随颜色主题），
        // 其次退回 html 上的 data-oe-theme，再退回系统偏好。
        const cls = document.body.classList;
        const dark = cls.contains('vscode-dark') || cls.contains('vscode-high-contrast') ||
            (document.documentElement.getAttribute('data-oe-theme') === 'dark');
        await renderMarkdownExtras(body, { theme: dark ? 'dark' : 'default' });
    } catch (err: any) {
        messages.push(`图表/公式渲染失败：${String(err?.message || err)}`);
    }
    setStatus(messages.length ? 'Markdown 预览（有清洗提示）' : 'Markdown 预览');
    if (messages.length) {
        const note = document.createElement('div');
        note.className = 'hint';
        note.textContent = messages.join('；');
        content.appendChild(note);
    }
}

async function renderText(content: HTMLElement, bytes: Uint8Array): Promise<void> {
    const { html } = await textToHtml(bytes);
    const body = document.createElement('div');
    body.className = 'txt-body';
    body.innerHTML = html;
    content.appendChild(body);

    setStatus('纯文本预览');
}

window.addEventListener('message', (e) => {
    const msg = e.data as HostToWebview;
    switch (msg.type) {
        case 'init':
            format = msg.format;
            mode = msg.mode || 'view';
            currentBytes = b64ToBytes(msg.bytes);
            applyBaseStyle();
            renderShell();
            pdfViewer?.setTitle(msg.title || '');
            zoom.fit(); // 新文档：缩放复位到适应宽度
            void renderContent();
            if (!readySent) {
                readySent = true;
                vscode.postMessage({ type: 'ready' });
            }
            break;
        case 'reverted':
            void renderContent();
            break;
        case 'setMode':
            // 标题栏的「查看 / 编辑」命令：PDF 切到编辑模式就是打开 overlay 工具层
            mode = msg.mode;
            void renderContent();
            break;
        case 'saved':
            pdfViewer?.markSaved();
            break;
        case 'info':
            break;
    }
});
