import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';
import { copyFileSync, mkdirSync } from 'node:fs';

/** 把 ppt-parser 的浏览器构建复制到 public/vendor，供 importmap 原生加载 */
function copyPptParserBrowser() {
    const src = fileURLToPath(new URL('./node_modules/@fefeding/ppt-parser/dist/ppt-parser.browser.js', import.meta.url));
    const destDir = fileURLToPath(new URL('./examples/public/vendor/', import.meta.url));
    mkdirSync(destDir, { recursive: true });
    copyFileSync(src, destDir + 'ppt-parser.browser.js');
    return {
        name: 'copy-ppt-parser-browser',
        buildStart() { copyFileSync(src, destDir + 'ppt-parser.browser.js'); }
    };
}

/** 本地示例页：Vite 直连 src 源码，改代码即时热更新 */
export default defineConfig({
    root: fileURLToPath(new URL('./examples', import.meta.url)),
    plugins: [copyPptParserBrowser()],
    resolve: {
        alias: {
            jdoc: fileURLToPath(new URL('./src/index.ts', import.meta.url))
        }
    },
    // ppt-parser 一律原生加载其浏览器构建（由 index.html 的 importmap 解析），
    // 不参与打包：经 bundler 打包后其媒体处理行为会改变（图片错页/丢失），
    // 导致与 ppt-parser 官方示例渲染不一致。
    optimizeDeps: { exclude: ['@fefeding/ppt-parser'] },
    server: {
        port: 5273,
        open: '/index.html'
    },
    build: {
        outDir: '../dist-examples',
        emptyOutDir: true,
        rollupOptions: { external: ['@fefeding/ppt-parser'] }
    }
});
