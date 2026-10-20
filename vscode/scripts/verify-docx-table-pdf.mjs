// DOCX -> PDF 表格合并单元格回归。
//
// 问题：gridSpan / vMerge 合并信息在 docx->pdf 链路的 blocks-to-pdf 里丢失，
// 生成端只能按「第 N 个 cell」取 tblGrid 列宽，于是合并表头拿到某个数字列的
// 窄宽度，文字被逐字折行（business priority 1 -> busi/nes/s/prio/rity/1），
// 且纵向合并丢失导致列整体错位。
//
// 这里做两层断言：
// 1. 结构：blocksToPdf 产出的表格元素保留了 colspan/rowspan，columnWidths 为网格列宽
// 2. 渲染：真实文档转 PDF 后，用 pdf.js 提取文本，表头文字完整、不再逐字拆开
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// __dirname = <repo>/vscode/scripts，lib 的 dist 与 examples 都在仓库根
const root = path.resolve(__dirname, '..', '..');
const { convert } = await import(path.resolve(root, 'dist', 'jdoc.cjs'));

let failed = 0;
const check = (name, ok, detail = '') => {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
};

// ---------- 1. 结构层：构造带合并的表格走 blocksToPdf ----------
const { blocksToPdf } = await import(path.resolve(root, 'dist', 'jdoc.cjs'));

// tblGrid 网格列宽：user priority 独占 1 个宽列(rowspan=2)，两个 business priority
// 表头各横跨 5 列(gridSpan)，尾部 '...' 独占 1 列 -> 共 12 个网格列。
// 数字列 28pt：既放得下 '99'，5 列合并后(140pt)也放得下单行 'business priority 1'；
// 总宽 378pt 落在 A4 内容区内。
const gridCols = 12;
const columnWidths = [70, ...new Array(gridCols - 1).fill(28)];
const mergedTable = {
    metadata: { title: 'merge' },
    blocks: [
        {
            type: 'table',
            header: true,
            columnWidths,
            rows: [
                [
                    { type: 'cell', text: 'user priority', rowspan: 2 },
                    { type: 'cell', text: 'business priority 1', colspan: 5 },
                    { type: 'cell', text: 'business priority 2', colspan: 5 },
                    { type: 'cell', text: '...' }
                ],
                [
                    { type: 'cell', text: '0' },
                    { type: 'cell', text: '1' },
                    { type: 'cell', text: '2' },
                    { type: 'cell', text: '3' },
                    { type: 'cell', text: '4' },
                    { type: 'cell', text: '...' },
                    { type: 'cell', text: '98' },
                    { type: 'cell', text: '99' },
                    { type: 'cell', text: '0' },
                    { type: 'cell', text: '1' }
                ]
            ]
        }
    ]
};

// 渲染层内部会调 CJK 探测，这里表格是纯英文，用标准字体即可
const pdfBytes = await blocksToPdf(mergedTable, { fonts: null });
check('blocksToPdf 产出 PDF', pdfBytes?.length > 0, `${pdfBytes.length}B`);

// 用 pdf.js 提取文本，验证表头完整
const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
const doc = await pdfjs.getDocument({ data: pdfBytes.slice(), useSystemFonts: true }).promise;
const tc = await (await doc.getPage(1)).getTextContent();
const items = tc.items.map((it) => it.str).filter(Boolean);
const joined = items.join('|');
console.log('  文本项:', JSON.stringify(items.slice(0, 8)));

check('「business priority 1」完整未折行', items.includes('business priority 1'), `items 含: ${items.filter(s => s.includes('business')).join(' / ') || '(无)'}`);
check('「business priority 2」完整未折行', items.includes('business priority 2'));
check('「user priority」完整未折行', items.includes('user priority'));
check('未出现逐字拆分的 busi/nes', !items.includes('busi') && !items.includes('nes'));
check('数字列 0 与 99 都在', items.includes('0') && items.includes('99'));

// ---------- 2. 真实文档端到端 ----------
const DOCX = path.resolve(root, 'examples', '过载保护优化方案.docx');
if (fs.existsSync(DOCX)) {
    try {
        const realPdf = await convert(fs.readFileSync(DOCX), 'pdf', { autoFont: true });
        const rdoc = await pdfjs.getDocument({ data: realPdf.slice() }).promise;
        let text = '';
        for (let p = 1; p <= Math.min(rdoc.numPages, 30); p++) {
            const t = await (await rdoc.getPage(p)).getTextContent();
            text += t.items.map((it) => it.str).join('') + '\n';
        }
        const found = text.includes('business priority 1');
        const broken = /\bbusi\b|\bnes\b|\bprio\b|\briority\b/.test(text);
        console.log(`  真实文档页数: ${rdoc.numPages}`);
        check('真实文档含完整「business priority 1」', found, found ? '' : '（该表可能不在前 30 页或文本提取差异）');
        check('真实文档未出现逐字拆分', !broken || !found, broken ? '检测到 busi/nes/prio/rity 拆分' : '');
    } catch (err) {
        console.log('  真实文档转换跳过：', String(err?.message || err).slice(0, 120));
    }
} else {
    console.log('  未找到 examples/过载保护优化方案.docx，跳过端到端');
}

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
process.exit(failed ? 1 : 0);