// 文档互转示例（@fefeding/jdoc 的 convert 技能附带资源）
// 用法：node example.mjs <源文件> <目标格式 docx|pdf|md|html>
import { readFile, writeFile } from 'node:fs/promises';
import { convert } from '@fefeding/jdoc';

const [src, target] = process.argv.slice(2);
if (!src || !target) {
  console.error('用法：node example.mjs <源文件> <docx|pdf|md|html>');
  process.exit(1);
}

const input = await readFile(src);
const out = await convert(input, target);

const dest = `out.${target}`;
await writeFile(dest, typeof out === 'string' ? out : out);
console.log(`已生成 ${dest}`);
