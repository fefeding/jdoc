import * as vscode from 'vscode';
import { JdocDocument } from './document';
import type { DocFormat, HostToWebview, WebviewToHost } from './protocol';

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

    constructor(private readonly context: vscode.ExtensionContext) {}

    async openCustomDocument(
        uri: vscode.Uri,
        _openContext: vscode.CustomDocumentOpenContext,
        _token: vscode.CancellationToken
    ): Promise<JdocDocument> {
        const document = await JdocDocument.create(uri);
        document.onDidChange(() => this._onDidChangeCustomDocument.fire({ document }));
        document.onDidRevert(() => this._onDidChangeCustomDocument.fire({ document }));
        document.onDidSave(() => this._onDidChangeCustomDocument.fire({ document }));
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

        webviewPanel.onDidChangeViewState((e) => {
            if (e.webviewPanel.active) this._activeDoc = document;
        });

        const send = (msg: HostToWebview) => webviewPanel.webview.postMessage(msg);
        const sendInit = async () => {
            const s = document.snapshot();
            send({
                type: 'init',
                format: s.format,
                mode: s.mode,
                bytes: s.bytes,
                title: document.uri.path.split('/').pop() || 'document',
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
