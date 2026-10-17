import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import json from '@rollup/plugin-json';
import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';
import dts from 'rollup-plugin-dts';
import fs from 'fs';
import path from 'path';

const pkg = JSON.parse(fs.readFileSync(path.resolve('./package.json'), 'utf-8'));

const banner = `/**
 * ${pkg.name} v${pkg.version}
 * ${pkg.description}
 * MIT License
 */`;

// 运行期全部保持 external：四种格式的解析/生成依赖体积都很大，
// 内联进产物会产出 10MB+ 的单文件，且升级依赖无法通过 semver 生效。
const external = [
  ...Object.keys(pkg.dependencies || {}),
  ...Object.keys(pkg.optionalDependencies || {}),
  ...Object.keys(pkg.peerDependencies || {}),
  // 仅在 Node 分支执行的动态探测用到这两个内置模块，
  // 显式声明可避免 browser 产物里出现未解析依赖告警
  /^node:/
];

const ts = () =>
  typescript({
    tsconfig: './tsconfig.json',
    compilerOptions: { checkJs: false, noEmitOnError: false, declaration: false }
  });

export default [
  // Node/通用版本：ESM + CJS 双格式，依赖保持 external
  {
    input: 'src/index.ts',
    output: [
      { file: './dist/jdoc.esm.js', format: 'es', banner, sourcemap: true, exports: 'named' },
      { file: './dist/jdoc.cjs', format: 'cjs', banner, sourcemap: true, exports: 'named' }
    ],
    plugins: [nodeResolve({ extensions: ['.ts', '.js', '.json'] }), commonjs(), json(), ts()],
    external,
    treeshake: false
  },
  // 浏览器版本：ESM 单文件，仍保持 external（pdfjs-dist 需要独立 worker 资源，
  // 无法内联；其余依赖由消费方的 bundler 处理）
  {
    input: 'src/index.ts',
    output: {
      file: './dist/jdoc.browser.js',
      format: 'es',
      banner,
      sourcemap: true,
      exports: 'named'
    },
    plugins: [
      nodeResolve({ browser: true, preferBuiltins: false, extensions: ['.ts', '.js', '.json'] }),
      commonjs({ include: /node_modules/ }),
      json(),
      ts()
    ],
    external,
    treeshake: false
  },
  // 类型声明：自包含 .d.ts
  {
    input: 'src/index.ts',
    output: [{ file: pkg.types, format: 'es' }],
    plugins: [dts()],
    external
  }
];