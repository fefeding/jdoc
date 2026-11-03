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

/**
 * 把 pdfjs 的 cMaps / 标准字体 / wasm 资源复制到 media/pdfjs/。
 *
 * 中文等 CID 字体的 PDF 必须有 cMaps 才能把 CID 解码成字形，缺了的表现是
 * 「线条、色块、项目符号都渲染出来，唯独文字全部消失」；standard_fonts 是
 * 未内嵌字体（Helvetica 等 14 种标准字体）的替代字形；wasm 用于 JPEG2000
 * 等图像解码。webview 端 getDocument 需显式传 cMapUrl / standardFontDataUrl /
 * wasmUrl（由 extension.ts 注入全局量，util.ts 组装）。
 */
function copyPdfAssets() {
  let base;
  try {
    const req = require('module').createRequire(path.join(__dirname, 'package.json'));
    base = path.dirname(req.resolve('pdfjs-dist/package.json'));
  } catch {
    base = path.join(__dirname, '..', 'node_modules', 'pdfjs-dist');
  }
  const outDir = path.join(__dirname, 'media', 'pdfjs');
  for (const sub of ['cmaps', 'standard_fonts', 'wasm']) {
    const from = path.join(base, sub);
    if (!fs.existsSync(from)) {
      console.warn(`[esbuild] 警告：未找到 pdfjs-dist/${sub}，相关 PDF 资源将缺失`);
      continue;
    }
    const to = path.join(outDir, sub);
    fs.cpSync(from, to, { recursive: true });
    console.log(`[esbuild] pdf ${sub} -> ${path.relative(__dirname, to)}`);
  }
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
  const assetsDir = path.join(mediaDir, 'pdfjs');
  // 原始 worker 取自 node_modules 的 pdfjs-dist（与打包进 webview.js 的版本一致），
  // 不能用 media 里那份——它可能已经是上一次构建的内嵌版本，否则会层层叠加。
  const pristine = path.join(__dirname, '..', 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.worker.min.mjs');
  if (!fs.existsSync(pristine) || !fs.existsSync(assetsDir)) {
    console.warn('[esbuild] 警告：缺少 pdfjs-dist worker 或 media/pdfjs，跳过资源内嵌');
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

  // DOCX / XLSX 编辑器：独立 webview 入口（查看 + 块式/网格编辑），同样自包含。
  const officeCt = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'webview', 'office-editor', 'main.ts')],
    outfile: path.join(__dirname, 'dist', 'office-editor.js'),
    format: 'iife',
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
    copyPdfAssets();
    embedPdfAssetsIntoWorker();
    await Promise.all([extCtx.watch(), webCtx.watch(), pptxCt.watch(), officeCt.watch()]);
    console.log('[esbuild] watching for changes… (press Ctrl+C to stop)');
    console.log('[esbuild] 调试：按 F5 选择 "Run Extension (watch)"；改代码后在扩展开发宿主窗口按 Ctrl/Cmd+R 刷新。');
    return;
  }

  await Promise.all([extCtx.rebuild(), webCtx.rebuild(), pptxCt.rebuild(), officeCt.rebuild()]);
  copyPdfWorker();
  copyPdfAssets();
  embedPdfAssetsIntoWorker();
  await extCtx.dispose();
  await webCtx.dispose();
  await pptxCt.dispose();
  await officeCt.dispose();
  console.log('[esbuild] build complete.');
}

buildWithContext().catch((e) => {
  console.error(e);
  process.exit(1);
});
