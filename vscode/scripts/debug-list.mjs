// 诊断 5：列表排版差异——文档里的真实编号/缩进/间距 vs 我们的渲染。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const jdoc = require(path.resolve(__dirname, '..', '..', 'dist', 'jdoc.cjs'));
const file = path.resolve(__dirname, '..', '..', 'examples', '过载保护优化方案.docx');
const doc = await jdoc.docxToJson(fs.readFileSync(file));

const lists = doc.blocks.filter((b) => b.type === 'list');
console.log(`列表块数: ${lists.length}`);
for (const [i, l] of lists.slice(0, 3).entries()) {
    console.log(`\n[列表 ${i + 1}] ordered=${l.ordered} items=${l.items.length}`);
    for (const it of l.items.slice(0, 4)) {
        const runs = typeof it === 'string' ? [] : it;
        const text = typeof it === 'string' ? it : runs.map((r) => r.text).join('');
        console.log(`   - ${JSON.stringify(text.slice(0, 40))} sizes=${JSON.stringify([...new Set(runs.map((r) => r.size).filter(Boolean))])}`);
    }
    console.log(`   字段: ${JSON.stringify(Object.keys(l))}`);
}

// numbering 定义：编号格式（decimal/lowerLetter…）
const zip = fs.readFileSync(file);
console.log('\n=== numbering 摘要 ===');
const xml = require('node:zlib');
void xml;