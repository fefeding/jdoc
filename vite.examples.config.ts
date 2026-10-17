import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

/** 本地示例页：Vite 直连 src 源码，改代码即时热更新 */
export default defineConfig({
    root: fileURLToPath(new URL('./examples', import.meta.url)),
    resolve: {
        alias: [
            { find: 'jdoc', replacement: fileURLToPath(new URL('./src/index.ts', import.meta.url)) },
            // 图表渲染依赖 ECharts：直接复用 ppt-parser 自带的构建（与 ppt-parser
            // 官方示例同源，无需额外依赖）。该包未导出 echarts 子路径，故用别名指向文件。
            { find: /^echarts-gl$/, replacement: fileURLToPath(new URL('./node_modules/@fefeding/ppt-parser/examples/chart-lib/echarts-gl.min.js', import.meta.url)) },
            { find: /^echarts$/, replacement: fileURLToPath(new URL('./node_modules/@fefeding/ppt-parser/examples/chart-lib/echarts.min.js', import.meta.url)) }
        ]
    },
    server: {
        port: 5273,
        open: '/index.html'
    },
    build: {
        outDir: '../dist-examples',
        emptyOutDir: true,
        // echarts 体积较大，单独成 chunk 避免主包过大
        rollupOptions: {
            output: {
                manualChunks(id) {
                    if (id.includes('echarts')) return 'echarts';
                    return undefined;
                }
            }
        }
    }
});
