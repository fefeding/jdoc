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

/**
 * 把 pdfjs 的 worker 复制到 dist/。
 *
 * 为什么必须单独放一个文件：pdfjs v5+ 的 fake worker 依赖
 * `GlobalWorkerOptions.workerSrc`，它源码里的内置默认值是相对路径
 * "./pdf.worker.mjs"——pdfjs 被 esbuild 内联进 extension.js 后，这个相对路径
 * 会指向 dist/ 目录本身，worker 必然找不到，宿主侧的 PDF 转换会整体失败
 * （报「No "GlobalWorkerOptions.workerSrc" specified」）。
 *
 * 放在 dist/ 里与 extension.js 同级，jdoc 的自动探测（同目录兜底）就能命中；
 * 同时 extension.ts 也会显式把它的 file:// 地址传给 convert()。
 */
function copyPdfWorker() {
  const outDir = path.join(__dirname, 'dist');
  const dest = path.join(outDir, 'pdf.worker.mjs');
  const candidates = [];
  try {
    // 从 vscode 目录解析 @fefeding/jdoc 的依赖链里的 pdfjs-dist
    const require = require('module').createRequire(path.join(__dirname, 'package.json'));
    candidates.push(path.join(path.dirname(require.resolve('pdfjs-dist/package.json')), 'legacy', 'build', 'pdf.worker.mjs'));
  } catch {
    // 解析失败时退回仓库根目录的 node_modules
  }
  candidates.push(
    path.join(__dirname, '..', 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.mjs')
  );
  for (const src of candidates) {
    try {
      if (!fs.existsSync(src)) continue;
      fs.mkdirSync(outDir, { recursive: true });
      fs.copyFileSync(src, dest);
      console.log(`[esbuild] pdf worker -> ${path.relative(__dirname, dest)}`);
      return;
    } catch {
      // 试下一个候选路径
    }
  }
  console.warn('[esbuild] 警告：未找到 pdf.worker.mjs，PDF 转换在宿主侧将不可用');
}

/** 内嵌 worker 文件的标记（webview 端据此判定资源已内嵌，无需再探 URL） */
const EMBEDDED_WORKER_MARKER = '/* JDOC_PDF_EMBEDDED_ASSETS */';

/**
 * 把 cMap / 标准字体 / wasm 以 base64 直接嵌进 webview 用的 pdf worker。
 *
 * 为什么：pdfjs 在 worker 内部 fetch(cMapUrl + 文件名)。VS Code 官方 webview 的
 * 页面 origin 与 asWebviewUri 资源 origin 不同，且 worker 又是 blob: URL —— 跨源
 * fetch 会被 CORS/scheme 拦掉（CodeBuddy 同源所以能过）。表现就是「CodeBuddy 正常、
 * VS Code 调试宿主白字」。这里把资源字节放进 worker 源码并覆盖其 fetch，让 worker
 * 完全不发网络请求，任何宿主都一致。
 */
function embedPdfAssetsIntoWorker() {
  const mediaDir = path.join(__dirname, 'media');
  const workerPath = path.join(mediaDir, 'pdf.worker.min.mjs');
  // 资源素材直接取自 node_modules/pdfjs-dist（构建时用、不进包），不再落盘到 media/pdfjs。
  // 内嵌后 webview 的 worker 自带全部 cMap/字体/wasm，media/pdfjs 目录因此可删，插件瘦身。
  let pdfjsDist;
  try {
    const req = require('module').createRequire(path.join(__dirname, 'package.json'));
    pdfjsDist = path.dirname(req.resolve('pdfjs-dist/package.json'));
  } catch {
    pdfjsDist = path.join(__dirname, '..', 'node_modules', 'pdfjs-dist');
  }
  const assetsDir = pdfjsDist;
  // 原始 worker 取自 node_modules 的 pdfjs-dist（与打包进 webview.js 的版本一致），
  // 不能用 media 里那份——它可能已经是上一次构建的内嵌版本，否则会层层叠加。
  const pristine = path.join(__dirname, '..', 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.min.mjs');
  if (!fs.existsSync(pristine) || !fs.existsSync(path.join(assetsDir, 'cmaps'))) {
    console.warn('[esbuild] 警告：缺少 pdfjs-dist worker 或 cmaps，跳过资源内嵌');
    return;
  }
  const dict = {};
  let bytes = 0;
  for (const sub of ['cmaps', 'standard_fonts', 'wasm']) {
    const dir = path.join(assetsDir, sub);
    if (!fs.existsSync(dir)) continue;
    for (const name of fs.readdirSync(dir)) {
      const file = path.join(dir, name);
      if (!fs.statSync(file).isFile()) continue;
      const buf = fs.readFileSync(file);
      dict[`${sub}/${name}`] = buf.toString('base64');
      bytes += buf.length;
    }
  }
  const interceptor = `(() => {
  const assets = globalThis.__JDC_PDF_ASSETS__ || {};
  const nativeFetch = globalThis.fetch ? globalThis.fetch.bind(globalThis) : null;
  if (!nativeFetch) return;
  globalThis.fetch = async function (input, init) {
    try {
      let url = typeof input === 'string' ? input : (input && typeof input.url === 'string' ? input.url : '');
      const q = url.indexOf('?'); if (q >= 0) url = url.slice(0, q);
      const h = url.indexOf('#'); if (h >= 0) url = url.slice(0, h);
      for (const dir of ['cmaps/', 'standard_fonts/', 'wasm/']) {
        const i = url.lastIndexOf('/' + dir);
        if (i < 0) continue;
        const b64 = assets[url.slice(i + 1)];
        if (!b64) break;
        const bin = atob(b64);
        const out = new Uint8Array(bin.length);
        for (let k = 0; k < bin.length; k++) out[k] = bin.charCodeAt(k);
        return new Response(out, { status: 200, headers: { 'Content-Type': 'application/octet-stream' } });
      }
    } catch (_) { /* 交给原生 fetch */ }
    return nativeFetch(input, init);
  };
})();`;
  const prelude = `${EMBEDDED_WORKER_MARKER}\nglobalThis.__JDC_PDF_ASSETS__=${JSON.stringify(dict)};\n${interceptor}\n`;
  fs.writeFileSync(workerPath, prelude + fs.readFileSync(pristine, 'utf8'));
  console.log(`[esbuild] pdf worker 内嵌 ${Object.keys(dict).length} 个资源 (${(bytes / 1048576).toFixed(1)}MB) -> ${path.relative(__dirname, workerPath)}`);
}

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

  // 三个浏览器端入口合并到同一上下文并开启 splitting：@fefeding/jdoc 等公共依赖
  // 会被抽成共享 chunk（dist/chunk-*.js），三份入口不再各自打包一份 jdoc，
  // 安装体积省约 4MB 且功能完全等价。入口改为 ESM 模块，由 extension.ts 以
  // <script type="module"> 加载；ECharts / echarts-gl 仍以全局 <script> 注入。
  const webSharedCtx = await esbuild.context({
    ...base,
    entryPoints: {
      webview: path.join(__dirname, 'src', 'webview', 'main.ts'),
      'pptx-editor': path.join(__dirname, 'src', 'webview', 'pptx-editor', 'main.ts'),
      'office-editor': path.join(__dirname, 'src', 'webview', 'office-editor', 'main.ts')
    },
    outdir: path.join(__dirname, 'dist'),
    format: 'esm',
    splitting: true,
    platform: 'browser',
    target: 'es2020',
    alias: {
      '@fefeding/jdoc': resolveJdocEntry()
    },
    plugins: [nodeStubPlugin]
  });

  if (watch) {
    // watch 不 rebuild，但静态资源（含内嵌 cMap 的 worker）必须先生成一次，
    // 否则 F5 "Run Extension (watch)" 会拿到旧 worker，VS Code 调试宿主仍白字。
    copyPdfWorker();
    embedPdfAssetsIntoWorker();
    await Promise.all([extCtx.watch(), webSharedCtx.watch()]);
    console.log('[esbuild] watching for changes… (press Ctrl+C to stop)');
    console.log('[esbuild] 调试：按 F5 选择 "Run Extension (watch)"；改代码后在扩展开发宿主窗口按 Ctrl/Cmd+R 刷新。');
    return;
  }

  await Promise.all([extCtx.rebuild(), webSharedCtx.rebuild()]);
  copyPdfWorker();
  embedPdfAssetsIntoWorker();
  await extCtx.dispose();
  await webSharedCtx.dispose();
  console.log('[esbuild] build complete.');
}

buildWithContext().catch((e) => {
  console.error(e);
  process.exit(1);
});
