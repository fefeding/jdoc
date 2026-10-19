import * as vscode from 'vscode';
import { JdocDocument } from './document';
import { convert, type ConvertTarget } from '@fefeding/jdoc';
import type { DocFormat, DocTabInfo, EditorMode, HostToWebview, WebviewToHost } from './protocol';

/** 可被转换命令识别的源格式（与扩展名对应） */
const SUPPORTED_EXT = /\.(docx|pdf|md|markdown|html|htm|txt)$/;
/** 各目标格式对应的展示名 / 扩展名 / 保存过滤器 */
const TARGET_META: Record<ConvertTarget, { label: string; ext: string; filter: Record<string, string[]> }> = {
    docx: { label: 'Word 文档 (.docx)', ext: 'docx', filter: { Word: ['docx'] } },
    pdf: { label: 'PDF 文档 (.pdf)', ext: 'pdf', filter: { PDF: ['pdf'] } },
    md: { label: 'Markdown (.md)', ext: 'md', filter: { Markdown: ['md', 'markdown'] } },
    html: { label: 'HTML 网页 (.html)', ext: 'html', filter: { HTML: ['html', 'htm'] } }
};
/** 文件扩展名 -> 源格式（传给 convert 的 from 参数） */
function extToFormat(path: string): ConvertTarget | 'txt' | undefined {
    if (/\.docx$/i.test(path)) return 'docx';
    if (/\.pdf$/i.test(path)) return 'pdf';
    if (/\.markdown?$/i.test(path)) return 'md';
    if (/\.html?$/i.test(path)) return 'html';
    if (/\.txt$/i.test(path)) return 'txt';
    return undefined;
}

const VIEW_TYPE = 'jdoc.documentEditor';

/**
 * 本地化回退：无语言包时 vscode.l10n.t 会原样返回 key，这里给出英文回退。
 */
function loc(key: string, fallback: string, ...args: string[]): string {
    const value = vscode.l10n.t(key, ...args);
    return !value || value === key ? fallback.replace('{0}', args[0] ?? '') : value;
}

function getNonce(): string {
    let text = '';
    const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (let i = 0; i < 32; i++) text += possible.charAt(Math.floor(Math.random() * possible.length));
    return text;
}

class JdocEditorProvider implements vscode.CustomEditorProvider<JdocDocument> {
    public static readonly viewType = VIEW_TYPE;

    private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<
        vscode.CustomDocumentEditEvent<JdocDocument> | vscode.CustomDocumentContentChangeEvent<JdocDocument>
    >();
    public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

    private _activeDoc: JdocDocument | undefined;

    /** 当前活动的 webview 面板（标题栏命令需要向它发消息）。 */
    private _activePanel: vscode.WebviewPanel | undefined;

    /**
     * 已打开的文档（key 为 uri.toString()）。用于编辑器标题栏的「切换文档」命令：
     * 所有自定义编辑器都经由本 provider 打开，因此这里能拿到全量列表。
     */
    private readonly _openDocs = new Map<string, DocTabInfo>();

    constructor(private readonly context: vscode.ExtensionContext) {}

    async openCustomDocument(
        uri: vscode.Uri,
        _openContext: vscode.CustomDocumentOpenContext,
        _token: vscode.CancellationToken
    ): Promise<JdocDocument> {
        const document = await JdocDocument.create(uri);
        document.onDidChange(() => {
            this._onDidChangeCustomDocument.fire({ document });
            this.syncDocTab(document);
        });
        document.onDidRevert(() => {
            this._onDidChangeCustomDocument.fire({ document });
            this.syncDocTab(document);
        });
        document.onDidSave(() => {
            this._onDidChangeCustomDocument.fire({ document });
            this.syncDocTab(document);
        });
        return document;
    }

    async resolveCustomEditor(
        document: JdocDocument,
        webviewPanel: vscode.WebviewPanel,
        _token: vscode.CancellationToken
    ): Promise<void> {
        webviewPanel.webview.options = {
            enableScripts: true,
            enableCommandUris: false,
            localResourceRoots: [this.context.extensionUri]
        };
        webviewPanel.webview.html = this.getWebviewHtml(webviewPanel.webview, document.format);

        const key = document.uri.toString();
        const title = document.uri.path.split('/').pop() || 'document';
        this._openDocs.set(key, { uri: key, title, format: document.format, dirty: document.isDirty });
        webviewPanel.onDidDispose(() => {
            this._openDocs.delete(key);
        });

        webviewPanel.onDidChangeViewState((e) => {
            if (e.webviewPanel.active) {
                this._activeDoc = document;
                this._activePanel = webviewPanel;
                this.syncContext();
            }
        });

        const send = (msg: HostToWebview) => webviewPanel.webview.postMessage(msg);
        const sendInit = async () => {
            const s = document.snapshot();
            send({
                type: 'init',
                format: s.format,
                mode: s.mode,
                bytes: s.bytes,
                title,
                uri: key,
                dirty: document.isDirty,
                locale: vscode.env.language
            });
        };

        document.onDidSave(() => send({ type: 'saved' }));
        document.onDidRevert(() => {
            void sendInit();
            send({ type: 'reverted' });
        });

        webviewPanel.webview.onDidReceiveMessage(async (message: WebviewToHost) => {
            try {
                switch (message.type) {
                    case 'ready':
                        await sendInit();
                        break;
                    case 'setMode':
                        document.mode = message.mode;
                        if (this._activeDoc === document) this.syncContext();
                        break;
                    case 'sync':
                        document.setBytes(message.bytes);
                        document.setDirty(message.dirty);
                        break;
                    case 'save':
                        await document.save();
                        break;
                    case 'saveAs':
                        await this.saveAs(document, message.bytes);
                        break;
                    case 'revert':
                        await document.revert();
                        await sendInit();
                        send({ type: 'info', kind: 'info', message: loc('extension.reloaded', 'Reloaded') });
                        break;
                    case 'error':
                        send({ type: 'info', kind: 'error', message: message.message });
                        break;
                }
            } catch (err: any) {
                send({ type: 'info', kind: 'error', message: err?.message || String(err) });
            }
        });

        try {
            await sendInit();
        } catch (err: any) {
            send({
                type: 'info',
                kind: 'error',
                message: loc('extension.initFailed', 'Initialization failed: {0}', String(err?.message || err))
            });
        }
    }

    /** 记录当前活动文档，用于标题栏命令（查看/编辑、切换文档）。 */
    private syncContext(): void {
        void vscode.commands.executeCommand('setContext', 'jdoc:editing', this._activeDoc?.mode === 'edit');
    }

    /** 标题栏「查看 / 编辑」命令：把模式推给当前活动的 webview。 */
    public async setActiveMode(mode: EditorMode): Promise<void> {
        const doc = this._activeDoc;
        if (!doc) return;
        doc.mode = mode;
        this.syncContext();
        await this._activePanel?.webview.postMessage({ type: 'setMode', mode } as HostToWebview);
    }

    /** 标题栏「切换文档」命令：在已打开的 jdoc 文档之间跳转。 */
    public async switchDocument(): Promise<void> {
        const docs = Array.from(this._openDocs.values());
        if (docs.length === 0) {
            vscode.window.showInformationMessage(loc('extension.noOpenDoc', 'No jdoc document is open.'));
            return;
        }
        const items = docs.map((d) => ({
            label: d.title,
            description: vscode.workspace.asRelativePath(vscode.Uri.parse(d.uri)),
            detail: d.dirty ? loc('extension.unsavedDoc', 'Unsaved changes') : undefined,
            uri: d.uri
        }));
        const picked = await vscode.window.showQuickPick(items, {
            placeHolder: loc('extension.switchDoc', 'Switch to an open document')
        });
        if (!picked) return;
        // 用同一个 viewType 打开：已打开则聚焦到该标签，未打开则新建。
        await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.parse(picked.uri), VIEW_TYPE);
    }

    /** 文档 dirty 状态变化时刷新已打开文档列表。 */
    private syncDocTab(document: JdocDocument): void {
        const rec = this._openDocs.get(document.uri.toString());
        if (!rec || rec.dirty === document.isDirty) return;
        rec.dirty = document.isDirty;
    }

    // ---- CustomEditorProvider 必需方法 ----
    async saveCustomDocument(document: JdocDocument, _token: vscode.CancellationToken): Promise<void> {
        await document.save();
    }
    async saveCustomDocumentAs(
        document: JdocDocument,
        target: vscode.Uri,
        _token: vscode.CancellationToken
    ): Promise<void> {
        await document.save(target);
    }
    async revertCustomDocument(document: JdocDocument, _token: vscode.CancellationToken): Promise<void> {
        await document.revert();
    }
    async backupCustomDocument(
        document: JdocDocument,
        context: vscode.CustomDocumentBackupContext,
        _token: vscode.CancellationToken
    ): Promise<vscode.CustomDocumentBackup> {
        await document.backup(context.destination);
        return {
            id: context.destination.toString(),
            delete: () => {
                vscode.workspace.fs.delete(context.destination).then(
                    () => undefined,
                    () => undefined
                );
            }
        };
    }

    private async saveAs(document: JdocDocument, bytes?: string | null): Promise<void> {
        const filters: Record<string, string[]> = {
            PowerPoint: ['pptx'],
            Word: ['docx'],
            Excel: ['xlsx'],
            PDF: ['pdf'],
            Markdown: ['md', 'markdown'],
            Text: ['txt', 'log']
        };
        const uri = await vscode.window.showSaveDialog({
            defaultUri: document.uri,
            saveLabel: loc('extension.exportCopy', 'Export Copy'),
            filters: { Document: filters[document.format] || ['*'] }
        });
        if (!uri) return;
        if (bytes) document.setBytes(bytes);
        await document.save(uri);
        vscode.window.showInformationMessage(loc('extension.exported', 'Exported copy: {0}', uri.fsPath));
    }

    public async revertActive(): Promise<void> {
        if (this._activeDoc) await this._activeDoc.revert();
    }
    public async exportActive(): Promise<void> {
        if (this._activeDoc) await this.saveAs(this._activeDoc);
    }

    /**
     * 文档格式转换入口：把选中的文件转换成另一种格式（docx / pdf / md / html 互转）。
     * 通过 QuickPick 选择目标格式，转换后在扩展宿主（Node）完成并另存为新文件。
     */
    public async convertUri(uri?: vscode.Uri): Promise<void> {
        const target = uri || vscode.window.activeTextEditor?.document.uri;
        if (!target) {
            vscode.window.showInformationMessage(loc('extension.pickConvert', 'Select a file to convert with jdoc.'));
            return;
        }
        const sourceName = target.path.split('/').pop() || 'document';
        const from = extToFormat(target.path);
        if (!from) {
            vscode.window.showWarningMessage(loc('extension.unsupported', 'Unsupported file format for conversion.'));
            return;
        }

        const items = (Object.keys(TARGET_META) as ConvertTarget[])
            .filter((k) => k !== from)
            .map((k) => ({ label: TARGET_META[k].label, target: k }));
        const picked = await vscode.window.showQuickPick(items, {
            placeHolder: loc('extension.convertTo', 'Convert "{0}" to which format?', sourceName),
            title: loc('extension.convertTitle', 'jdoc: 转换文档格式')
        });
        if (!picked) return;

        const meta = TARGET_META[picked.target];
        await vscode.window.withProgress(
            { location: vscode.ProgressLocation.Notification, title: loc('extension.converting', 'Converting document…') },
            async () => {
                try {
                    const data = await vscode.workspace.fs.readFile(target);
                    const result = await convert(data, picked.target, { from });
                    const base = sourceName.replace(/\.[^.]+$/, '');
                    const dir = target.path.replace(/[^/]+$/, '');
                    const outUri = await vscode.window.showSaveDialog({
                        defaultUri: target.with({ path: dir + base + '.' + meta.ext }),
                        saveLabel: loc('extension.saveResult', 'Save converted file'),
                        filters: meta.filter
                    });
                    if (!outUri) return;
                    const bytes = typeof result === 'string' ? new TextEncoder().encode(result) : result;
                    await vscode.workspace.fs.writeFile(outUri, bytes);
                    vscode.window.showInformationMessage(
                        loc('extension.converted', 'Converted and saved: {0}', outUri.fsPath)
                    );
                } catch (err: any) {
                    vscode.window.showErrorMessage(loc('extension.convertFailed', 'Conversion failed: {0}', String(err?.message || err)));
                }
            }
        );
    }

    private getWebviewHtml(webview: vscode.Webview, format: DocFormat): string {
        // pptx -> 幻灯片编辑器；docx/xlsx -> office 编辑器；pdf/md/txt -> 只读查看器
        const scriptFile =
            format === 'pptx'
                ? 'pptx-editor.js'
                : format === 'docx' || format === 'xlsx'
                    ? 'office-editor.js'
                    : 'webview.js';
        const scriptUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'dist', scriptFile)
        );
        // 图表依赖：echarts / echarts-gl 作为全局脚本按序加载（必须先于 webview.js）
        const echartsUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts.min.js')
        );
        const echartsGlUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'echarts-gl.min.js')
        );
        // PDF（pdfjs）worker：webview 是单一打包产物，import.meta.url 派生的 worker 地址无效，
        // 因此这里显式给出可访问的 worker 资源 URL，由 webview 端写入全局供 jdoc 复用。
        const pdfWorkerUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'pdf.worker.min.mjs')
        );
        // KaTeX 样式（Markdown 公式渲染），CSS 内部以相对路径引用 media/fonts/*.woff2
        const katexCssUri = webview.asWebviewUri(
            vscode.Uri.joinPath(this.context.extensionUri, 'media', 'katex.min.css')
        );
        const nonce = getNonce();
        const oeTheme =
            vscode.window.activeColorTheme.kind === vscode.ColorThemeKind.Dark ||
            vscode.window.activeColorTheme.kind === vscode.ColorThemeKind.HighContrast
                ? 'dark'
                : 'light';
        const csp = [
            `default-src 'none';`,
            `img-src ${webview.cspSource} data: blob:;`,
            `media-src ${webview.cspSource} data: blob:;`,
            // KaTeX 字体（Markdown 公式）
            `font-src ${webview.cspSource} data:;`,
            `style-src ${webview.cspSource} 'unsafe-inline';`,
            // echarts-gl 内部用 new Function() 解析 expr()，需 unsafe-eval 否则 3D 图表抛错；
            // pdfjs worker 以 Worker 形式加载，需 worker-src（含 blob: 兜底）
            `script-src 'nonce-${nonce}' 'unsafe-eval' ${webview.cspSource};`,
            `worker-src ${webview.cspSource} blob: data:;`
        ].join(' ');
        return `<!DOCTYPE html>
<html lang="${vscode.env.language}" data-oe-theme="${oeTheme}">
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Security-Policy" content="${csp}" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<link rel="stylesheet" href="${katexCssUri}" />
<title>jdoc Document</title>
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}">window.__JDC_PDF_WORKER__ = "${pdfWorkerUri}";</script>
<script nonce="${nonce}" src="${echartsUri}"></script>
<script nonce="${nonce}" src="${echartsGlUri}"></script>
<script nonce="${nonce}" src="${scriptUri}"></script>
</body>
</html>`;
    }
}

export function activate(context: vscode.ExtensionContext): void {
    const provider = new JdocEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
            webviewOptions: { retainContextWhenHidden: true },
            supportsMultipleEditorsPerDocument: false
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand('jdoc.revert', () => provider.revertActive()),
        vscode.commands.registerCommand('jdoc.exportCopy', () => provider.exportActive()),
        vscode.commands.registerCommand('jdoc.convert', (uri?: vscode.Uri) => provider.convertUri(uri)),
        // 标题栏（editor/title）里的查看/编辑与文档切换
        vscode.commands.registerCommand('jdoc.viewMode', () => provider.setActiveMode('view')),
        vscode.commands.registerCommand('jdoc.editMode', () => provider.setActiveMode('edit')),
        vscode.commands.registerCommand('jdoc.switchDocument', () => provider.switchDocument()),
        // Markdown / TXT 等文件 VS Code 有内置编辑器（内置优先级更高），
        // 因此提供显式入口，用 jdoc 预览当前文件或资源管理器里选中的文件。
        vscode.commands.registerCommand('jdoc.preview', async (uri?: vscode.Uri) => {
            const target = uri || vscode.window.activeTextEditor?.document.uri;
            if (!target) {
                vscode.window.showInformationMessage(loc('extension.pickFile', 'Select a file to preview with jdoc.'));
                return;
            }
            await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
        })
    );
}

export function deactivate(): void {
    // no-op
}
