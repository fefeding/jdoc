// 诊断 6：精确定位「1.背景」下 a/b/c/d 列表段落的 pPr（编号 + 缩进 + 间距）。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
void require;
const file = path.resolve(__dirname, '..', '..', 'examples', '过载保护优化方案.docx');
const zip = await JSZip.loadAsync(fs.readFileSync(file));
const xml = await zip.file('word/document.xml').async('string');
const numbering = await zip.file('word/numbering.xml').async('string');

// 取前 2 个带 numPr 的段落，打印其 pPr
const paras = xml.split('<w:p>').slice(1);
let shown = 0;
for (const p of paras) {
    if (!p.includes('<w:numPr>')) continue;
    const pPr = p.slice(0, p.indexOf('</w:pPr>') + 8);
    console.log(`--- 列表段落 #${++shown} ---`);
    console.log(pPr.replace(/></g, '>\n<').split('\n').filter((l) => /numPr|numId|ilvl|ind|spacing|pStyle|jc|numFmt|lvlText|lvlJc|pPr/.test(l)).join('\n'));
    const text = (p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('').slice(0, 30);
    console.log(`  文本: ${JSON.stringify(text)}\n`);
    if (shown >= 2) break;
}

// 找到这些 numId 对应的 abstractNum 的 lvl 0 定义
const numIds = [...xml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]).slice(0, 3);
console.log('=== 出现的前几个 numId:', numIds.join(', '));
for (const numId of numIds) {
    const numRe = new RegExp(`<w:num w:numId="${numId}"[^>]*>\\s*<w:abstractNumId w:val="(\\d+)"`);
    const abs = xml && numbering.match(numRe);
    const abstractId = abs ? abs[1] : null;
    if (!abstractId) continue;
    const lvlRe = new RegExp(`<w:abstractNum w:abstractNumId="${abstractId}"[\\s\\S]*?<w:lvl w:ilvl="0"[\\s\\S]*?</w:lvl>`);
    const lvl = numbering.match(lvlRe);
    if (!lvl) continue;
    const f = lvl[0].match(/<w:numFmt w:val="([^"]+)"/);
    const t = lvl[0].match(/<w:lvlText w:val="([^"]*)"/);
    const ind = lvl[0].match(/<w:ind([^/]*)\/>/);
    const spc = lvl[0].match(/<w:spacing([^/]*)\/>/);
    console.log(`numId=${numId} abstract=${abstractId} fmt=${f?.[1]} text=${JSON.stringify(t?.[1])} ind=${ind?.[1]?.trim()} spacing=${spc?.[1]?.trim()}`);
}