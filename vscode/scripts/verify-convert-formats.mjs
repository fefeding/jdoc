// 验证：宿主侧「转换格式」支持的源格式与目标格式
//
// 背景：Excel（.xlsx）曾在插件里点「转换格式」直接报「Unsupported file format for conversion」，
// 因为宿主只认 docx/pdf/md/html/txt。这里用与 extension.ts 完全相同的调用方式
// （convert(data, target, { from, workerSrc })）跑一遍所有源格式 × 目标格式，
// 并用 src/convert-formats.ts 的映射断言「哪些扩展名可以被识别」。
//
// 用法：node scripts/verify-convert-formats.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { build } from 'esbuild';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const jdoc = require(path.resolve(__dirname, '..', '..', 'dist', 'jdoc.cjs'));

const WORKER_SRC = pathToFileURL(path.resolve(__dirname, '..', 'dist', 'pdf.worker.mjs')).href;
const EXAMPLES = path.resolve(__dirname, '..', '..', 'examples');

let failed = 0;
function check(name, ok, detail = '') {
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  ' + detail : ''}`);
    if (!ok) failed++;
}

/** 编译 src/convert-formats.ts 并加载（该文件不依赖 vscode API，可独立引入） */
async function loadFormatModule() {
    const result = await build({
        entryPoints: [path.resolve(__dirname, '..', 'src', 'convert-formats.ts')],
        bundle: true,
        format: 'esm',
        platform: 'node',
        write: false,
        logLevel: 'silent'
    });
    const code = result.outputFiles[0].text;
    const url = 'data:text/javascript;base64,' + Buffer.from(code).toString('base64');
    return import(url);
}

const { extToFormat, targetFormats } = await loadFormatModule();

// ---- 1. 扩展名映射 ----
console.log('\n== 源格式识别（扩展名 -> from）==');
const mapping = [
    ['a.docx', 'docx'],
    ['a.pdf', 'pdf'],
    ['a.md', 'md'],
    ['a.markdown', 'md'],
    ['a.html', 'html'],
    ['a.htm', 'html'],
    ['a.txt', 'txt'],
    ['a.xlsx', 'xlsx'],
    ['/dir/销售 表.XLSX', 'xlsx'],
    ['a.pptx', 'pptx'],
    ['a.zip', undefined]
];
for (const [file, expected] of mapping) {
    const got = extToFormat(file);
    check(`${file} -> ${expected ?? '(不支持)'}`, got === expected, `got=${got}`);
}

const targets = ['docx', 'pdf', 'md', 'html'];
check('xlsx 的目标格式不含自身', !targetFormats('xlsx', targets).includes('xlsx'));
check('xlsx 的目标格式为 docx/pdf/md/html', targetFormats('xlsx', targets).join(',') === 'docx,pdf,md,html');
check('md 的目标格式去掉 md', !targetFormats('md', targets).includes('md'));

// ---- 2. 宿主侧真实转换 ----
console.log('\n== 宿主侧转换（真实样例）==');
const samples = [
    { file: '过载保护优化方案.docx', from: 'docx' },
    { file: '销售分析表 副本.xlsx', from: 'xlsx' },
    { file: '最新【人教版】9年级数学课本•上册.pdf', from: 'pdf' }
];

// 用库现造一个最小 PPTX，覆盖 pptx 源格式（examples 里没有现成 pptx 样例）
async function buildPptxSample() {
    const deck = {
        version: '1.0',
        slideSize: { width: 960, height: 540 },
        metadata: { title: '示例', author: 'verify' },
        slides: [
            {
                elements: [
                    { type: 'text', x: 0, y: 0, width: 960, height: 80, fontSize: 32,
                      paragraphs: [{ text: '季度汇报', runs: [{ text: '季度汇报', bold: true, fontSize: 32 }] }] },
                    { type: 'text', x: 0, y: 100, width: 960, height: 300,
                      paragraphs: [
                          { text: '完成目标', bullet: true, runs: [{ text: '完成目标' }] },
                          { text: '存在问题', bullet: true, runs: [{ text: '存在问题' }] }
                      ] }
                ]
            },
            {
                elements: [
                    { type: 'text', x: 0, y: 0, width: 960, height: 80, fontSize: 28,
                      paragraphs: [{ text: '数据', runs: [{ text: '数据', bold: true, fontSize: 28 }] }] },
                    { type: 'table', x: 0, y: 120, width: 400, height: 120,
                      rows: [{ cells: [{ text: '合并标题', colSpan: 2 }] },
                             { cells: [{ text: 'A' }, { text: 'B' }] },
                             { cells: [{ text: 'C' }, { text: 'D' }] }] },
                    { type: 'chart', x: 0, y: 260, width: 400, height: 240, title: '销量',
                      chartType: 'barChart', categories: ['Q1', 'Q2'], series: [{ name: '销量', values: [10, 20] }] }
                ]
            }
        ]
    };
    return jdoc.jsonToPptx(deck);
}
samples.push({ data: await buildPptxSample(), from: 'pptx', label: 'generated.pptx' });

for (const sample of samples) {
    let data;
    if (sample.data) {
        data = sample.data;
    } else {
        const full = path.join(EXAMPLES, sample.file);
        if (!fs.existsSync(full)) {
            check(`${sample.file} 样例存在`, false, full);
            continue;
        }
        data = new Uint8Array(fs.readFileSync(full));
    }
    for (const to of targetFormats(sample.from, targets)) {
        const started = Date.now();
        try {
            const result = await jdoc.convert(data, to, { from: sample.from, workerSrc: WORKER_SRC });
            const size = typeof result === 'string' ? Buffer.byteLength(result) : result.byteLength;
            // 扫描版 PDF 解析不出文本，产出 0 字节属于预期，故只断言「不抛错」
            check(
                `${sample.from} -> ${to}`,
                true,
                `${size} bytes, ${Date.now() - started}ms${size ? '' : '（源文档无可提取内容）'}`
            );
        } catch (error) {
            check(`${sample.from} -> ${to}`, false, String(error?.message || error));
        }
    }
}

// Excel 的 Markdown 结果应包含表名与表格管道行
{
    const full = path.join(EXAMPLES, '销售分析表 副本.xlsx');
    if (fs.existsSync(full)) {
        const data = new Uint8Array(fs.readFileSync(full));
        const md = await jdoc.convert(data, 'md', { from: 'xlsx' });
        check('xlsx -> md 含表名标题', /^#\s+\S+/m.test(md), md.slice(0, 40));
        check('xlsx -> md 含 GFM 表格行', /\|\s*-{3,}\s*\|/.test(md));
    }
}

// PPTX 的 Markdown 结果应包含幻灯片标题与项目符号列表
{
    const data = await buildPptxSample();
    const md = await jdoc.convert(data, 'md', { from: 'pptx' });
    check('pptx -> md 含幻灯片标题', /^#\s+季度汇报/m.test(md), md.slice(0, 40));
    check('pptx -> md 含项目符号列表', /^- 完成目标/m.test(md));
}

console.log(failed ? `\nFAILED (${failed} failure(s))` : '\nALL PASSED (0 failure(s))');
process.exit(failed ? 1 : 0);
