import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

/** 本地示例页：Vite 直连 src 源码，改代码即时热更新 */
export default defineConfig({
    root: fileURLToPath(new URL('./examples', import.meta.url)),
    resolve: {
        alias: {
            jdoc: fileURLToPath(new URL('./src/index.ts', import.meta.url))
        }
    },
    server: {
        port: 5273,
        open: '/index.html'
    },
    build: {
        outDir: '../dist-examples',
        emptyOutDir: true
    }
});
