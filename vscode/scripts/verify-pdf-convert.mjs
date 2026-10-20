// 回归验证：宿主侧 PDF 转换（PDF -> Markdown）链路可用。
//
// 为什么要有这个脚本：PDF 解析依赖 pdfjs 的 worker，而 pdfjs 被打包器内联进
// extension.js 后，它内置的相对 worker 路径（"./pdf.worker.mjs"）会指向 bundle
// 所在目录，worker 必然加载不到——PDF 转换整体失败，报
//「pdf 解析失败：No "GlobalWorkerOptions.workerSrc" specified」。
//
// 这个问题只在**打包后的真实产物**里出现：单独 require jdoc 库是正常的，
// webview 侧也有自己的 worker 方案（window.__JDC_PDF_WORKER__）不受影响。
// 所以必须在真实产物上验证，普通单元测试覆盖不到。
import { createRequire } from 'node:module';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const vscodeDir = path.resolve(here, '..');
const repoRoot = path.resolve(vscodeDir, '..');

let failures = 0;
function check(name, ok, detail = '') {
    if (ok) {
        console.log(`  PASS  ${name}${detail ? '  ' + detail : ''}`);
    } else {
        failures++;
        console.log(`  FAIL  ${name}  ${detail}`);
    }
}

console.log('== 宿主侧 PDF 转换（走真实产物 dist/extension.js）==');

// 1) worker 文件必须随产物一起分发
const workerPath = path.join(vscodeDir, 'dist', 'pdf.worker.mjs');
check('dist/pdf.worker.mjs 存在（esbuild 需复制）', existsSync(workerPath),
    workerPath.replace(repoRoot, '.'));

// 2) 加载真实产物，用桩替换 vscode 模块
const Module = require('node:module');
const origLoad = Module._load;
const cmds = new Map();
const noop = () => undefined;
const errors = [];
let written = null;

function mkUri(s) {
    return {
        toString: () => s,
        fsPath: s.replace('file://', ''),
        path: s.replace('file://', ''),
        scheme: 'file',
        with: () => mkUri(s)
    };
}

const vscodeStub = new Proxy({
    commands: {
        registerCommand: (id, cb) => { cmds.set(id, cb); return { dispose: noop }; },
        executeCommand: async () => undefined
    },
    window: {
        showQuickPick: async () => ({ target: 'md' }),
        showSaveDialog: async () => mkUri('file:///tmp/jdoc-out.md'),
        showInformationMessage: noop,
        showWarningMessage: noop,
        showErrorMessage: (m) => { errors.push(String(m)); },
        withProgress: async (_o, task) => task(),
        createStatusBarItem: () => ({ show: noop, hide: noop, dispose: noop }),
        registerCustomEditorProvider: () => ({ dispose: noop }),
        activeTextEditor: undefined
    },
    workspace: {
        fs: {
            readFile: async () => pdfBytes,
            writeFile: async (_u, b) => { written = Buffer.from(b).toString('utf8'); }
        },
        asRelativePath: (u) => String(u),
        onDidChangeConfiguration: () => ({ dispose: noop })
    },
    Uri: { parse: (s) => mkUri(s), file: (p) => mkUri('file://' + p) },
    env: { language: 'zh-cn' },
    l10n: { t: (key, ...a) => (a.length ? key.replace('{0}', String(a[0])) : key) },
    ProgressLocation: { Notification: 15, Window: 10 },
    extensions: { getExtension: () => undefined },
    EventEmitter: class { constructor() { this.event = noop; } fire() { } },
    ViewColumn: { Active: -1, Beside: -2, One: 1 },
    FileType: { File: 1, Directory: 2 },
    FileSystemError: { FileNotFound: class extends Error { } },
    Disposable: class { dispose() { } },
    CustomEditorProvider: class { },
    ExtensionMode: { Production: 1 },
    ThemeColor: class { }
}, { get: (t, k) => (k in t ? t[k] : noop) });

Module._load = function (req) {
    if (req === 'vscode') return vscodeStub;
    return origLoad.apply(this, arguments);
};

// 3) 生成一个含真实文字的小 PDF（用仓库里的 jdoc 构建产物）
const jdocCjs = require(path.join(repoRoot, 'dist', 'jdoc.cjs'));
const pdfBytes = Buffer.from(await jdocCjs.jsonToPdf({
    elements: [{ type: 'text', text: 'Hello PDF Convert' }]
}));

const ext = require(path.join(vscodeDir, 'dist', 'extension.js'));
await ext.activate({
    subscriptions: [],
    extensionUri: { fsPath: '/tmp/ext', toString: () => 'file:///tmp/ext' },
    extensionPath: '/tmp/ext',
    globalState: { get: () => undefined, update: async () => { } },
    workspaceState: { get: () => undefined, update: async () => { } },
    extensionMode: 1
});

const convertCmd = cmds.get('jdoc.convert');
check('jdoc.convert 命令已注册', !!convertCmd);

if (convertCmd) {
    await convertCmd(mkUri('file:///tmp/jdoc-test.pdf'));
    check('转换过程没有报错', errors.length === 0, errors[0] ? errors[0].slice(0, 160) : '');
    check('转换写出了文件', written !== null);
    check('Markdown 内容非空', !!written && written.trim().length > 0,
        written ? `长度=${written.trim().length}` : '');
    check('Markdown 含原文文字', !!written && written.includes('Hello PDF Convert'),
        written ? JSON.stringify(written.trim().slice(0, 40)) : '');
}

console.log(failures === 0 ? '\nALL PASSED (0 failure(s))' : `\nFAILED (${failures} failure(s))`);
process.exit(failures === 0 ? 0 : 1);
