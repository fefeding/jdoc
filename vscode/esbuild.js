// 构建脚本：打包扩展宿主（Node）与 Webview（浏览器）两个产物。
//
// 用法：
//   node esbuild.js              单次构建（npm run compile）
//   node esbuild.js --watch      持续构建（npm run dev / npm run watch）
//   node esbuild.js --production 生产构建（npm run vscode:prepublish）
//
// @fefeding/jdoc 优先用同级仓库根目录已构建的 dist/jdoc.browser.js（开发/CI 联调），
// 否则回退到 node_modules 中发布的包（通过 install:jdoc 安装）。
const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');

function resolveJdocEntry() {
  const local = path.resolve(__dirname, '..', 'dist', 'jdoc.browser.js');
  if (fs.existsSync(local)) return local;
  const fromNodeModules = path.resolve(__dirname, 'node_modules', '@fefeding', 'jdoc', 'dist', 'jdoc.browser.js');
  if (!fs.existsSync(fromNodeModules)) {
    throw new Error(
      'Cannot resolve @fefeding/jdoc browser build. Build jdoc first (npm run build at repo root) ' +
      'or run "npm run install:jdoc".'
    );
  }
  return fromNodeModules;
}

// Webview 运行在浏览器，jdoc 的浏览器构建里仍有少量「写盘」专属的动态 import
// （node:fs/promises / node:path，例如 writeDocumentBundle），webview 不会走到这些
// 分支；这里用安全的浏览器 stub 替代，使打包通过且不影响解析/渲染路径。
const NODE_STUB = `
const noop = () => { throw new Error('node builtin unavailable in webview'); };
export const mkdir = noop;
export const writeFile = noop;
export const readFile = noop;
export const stat = noop;
export const rm = noop;
export const readdir = noop;
export const existsSync = noop;
export const readFileSync = noop;
export const writeFileSync = noop;
export const join = (...a) => a.join('/');
export const dirname = (p) => String(p).split('/').slice(0, -1).join('/');
export const resolve = (...a) => a.join('/');
export const basename = (p) => String(p).split('/').pop() || '';
export const extname = (p) => { const parts = String(p).split('.'); return parts.length > 1 ? '.' + parts.pop() : ''; };
export const sep = '/';
export const relative = noop;
export const isAbsolute = () => false;
export const parse = () => ({ root: '', dir: '', base: '', ext: '', name: '' });
export const format = noop;
export default { join, dirname, resolve, basename, extname, sep, mkdir, writeFile, readFile, stat };
`;
const nodeStubPlugin = {
    name: 'node-stub',
    setup(build) {
        build.onResolve({ filter: /^node:/ }, (args) => ({ path: args.path, namespace: 'node-stub' }));
        build.onLoad({ filter: /.*/, namespace: 'node-stub' }, () => ({ contents: NODE_STUB, loader: 'js' }));
    }
};

const watch = process.argv.includes('--watch');
const production = process.argv.includes('--production');

async function buildWithContext() {
  const base = {
    bundle: true,
    sourcemap: !production,
    minify: production,
    logLevel: 'info',
    legalComments: 'none'
  };

  const extCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'extension.ts')],
    outfile: path.join(__dirname, 'dist', 'extension.js'),
    external: ['vscode'],
    format: 'cjs',
    platform: 'node',
    target: 'node16'
  });

  const webCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'webview', 'main.ts')],
    outfile: path.join(__dirname, 'dist', 'webview.js'),
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    // 仅别名 jdoc；图表渲染器 chart-renderer.js 已复制到 src/webview 下直接打包，
    // ECharts / echarts-gl 通过 media/ 以全局 <script> 注入（先于 webview.js 加载）。
    alias: {
      '@fefeding/jdoc': resolveJdocEntry()
    },
    plugins: [nodeStubPlugin]
  });

  // PPTX 编辑器：独立 webview 入口（自包含，保留 vscode-pptx 的自启动 + ready 逻辑），
  // 避免与查看版 webview.js 的双 message 监听 / 双 acquireVsCodeApi 冲突。
  const pptxCt = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'webview', 'pptx-editor', 'main.ts')],
    outfile: path.join(__dirname, 'dist', 'pptx-editor.js'),
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    alias: {
      '@fefeding/jdoc': resolveJdocEntry()
    },
    plugins: [nodeStubPlugin]
  });

  if (watch) {
    await Promise.all([extCtx.watch(), webCtx.watch(), pptxCt.watch()]);
    console.log('[esbuild] watching for changes… (press Ctrl+C to stop)');
    console.log('[esbuild] 调试：按 F5 选择 "Run Extension (watch)"；改代码后在扩展开发宿主窗口按 Ctrl/Cmd+R 刷新。');
    return;
  }

  await Promise.all([extCtx.rebuild(), webCtx.rebuild(), pptxCt.rebuild()]);
  await extCtx.dispose();
  await webCtx.dispose();
  await pptxCt.dispose();
  console.log('[esbuild] build complete.');
}

buildWithContext().catch((e) => {
  console.error(e);
  process.exit(1);
});
