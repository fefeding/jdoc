# @fefeding/jdoc

A unified document parsing and generation library for **PPTX / DOCX / XLSX / PDF** — one API shape, works in both browsers and Node.js.

English | [简体中文](./README.zh-CN.md)

## Features

- **One naming convention, four formats** — parse as `xxxToYyy(input, options)`, generate as `jsonToXxx(doc, options)`
- **Zero-friction input** — pass `ArrayBuffer`, TypedArray, `Buffer`, `Blob`, `File`, base64 or a data URL directly
- **Uniform output** — every generator returns `Uint8Array`, ready for `new Blob([bytes])` or `fs.writeFile`
- **JSON as the contract** — parse results can be fed straight back into the generators, so parse → edit → generate round-trips work
- **Built-in HTML rendering** — DOCX and XLSX results render to inline-styled HTML out of the box
- **No framework dependency** — pure TypeScript shipping ESM, CJS and browser bundles

## Rendered output

Both screenshots come from the [`jdoc` VS Code extension](./vscode), which is built entirely on the APIs below — the paginated DOCX view (with table-of-contents sidebar) and the Excel-like XLSX grid (sheet tabs, formula bar, live formulas and charts) are rendered from the parsed JSON:

| DOCX | XLSX |
| --- | --- |
| ![DOCX](./vscode/images/docx.png) | ![XLSX](./vscode/images/xls.png) |

## Install

```bash
npm install @fefeding/jdoc
```

## Quick start

### Parse

```javascript
import { docxToHtml, xlsxToJson, pptxToStandard, pdfToText, detectFormat } from '@fefeding/jdoc';

const html = await docxToHtml(file);      // DOCX -> HTML
const { sheets } = await xlsxToJson(file); // XLSX -> workbook
const doc = await pptxToStandard(file);   // PPTX -> standard document
const text = await pdfToText(file);       // PDF -> plain text

console.log(await detectFormat(file));    // 'docx' | 'xlsx' | 'pptx' | 'pdf'
```

### Generate

```javascript
import { jsonToDocx, jsonToXlsx, jsonToPptx, jsonToPdf } from '@fefeding/jdoc';

const docx = await jsonToDocx([
  { type: 'heading', level: 1, text: 'Weekly Report' },
  { type: 'paragraph', text: 'Shipped modules A and B.' },
  { type: 'table', header: true, rows: [['Module', 'Progress'], ['A', '100%']] }
]);

const xlsx = await jsonToXlsx([{ name: 'Orders', rows: [['Name', 'Amount'], ['Alice', 1200]] }]);

const pdf = await jsonToPdf({
  metadata: { title: 'Report' },
  elements: [
    { type: 'heading', level: 1, text: 'Quarterly Report' },
    { type: 'text', text: 'Body copy…' },
    { type: 'table', header: true, rows: [['Item', 'Value'], ['Revenue', '1200']] }
  ]
});

const pptx = await jsonToPptx({
  slides: [{ elements: [{ type: 'text', text: 'Hello', x: 50, y: 50, width: 200, height: 40 }] }]
});
```

## API

### Shared

| Function | Description |
| --- | --- |
| `detectFormat(input)` | Identify the format from magic bytes and ZIP internals |
| `JdocError` | Unified error type carrying `format`, `code` and `cause` |
| `toUint8Array` / `toArrayBuffer` / `bytesToBase64` | Input and encoding helpers |
| `parseXml` / `findAll` / `findFirst` / `attr` / `textOf` | Zero-dependency XML utilities |

### DOCX

`docxToJson(input, options?)` → `{ metadata, blocks, messages, text }`
`docxToHtml(input, options?)` → the same plus `html`
`docxToText(input)` → `string`
`jsonToDocx(doc | blocks, options?)` → `Uint8Array`
`blocksToHtml(doc | blocks, options?)` → `string`

Block types: `heading`, `paragraph`, `list`, `table`, `image`, `quote`, `code`, `hr`, `pageBreak`.
Run styles: `bold`, `italic`, `underline`, `strike`, `subscript`, `superscript`, `color`, `size`, `font`, `highlight`, `link`.

### XLSX

`xlsxToJson(input, options?)` → `{ metadata, sheets, messages }`
`xlsxToHtml(input, options?)` → the same plus `html`
`xlsxToMatrix(input, sheetName?)` → `{ headers, rows }`, ready for charting libraries
`jsonToXlsx(workbook | sheets, options?)` → `Uint8Array`
`sheetsToHtml(workbook | sheets, options?)` → `string`

```javascript
rows: [
  ['plain value'],
  [{ 1: 'keyed by column index' }, { C2: 'or by A1 reference' }],
  [{ value: 1200, style: { numFmt: '#,##0.00', align: 'right' } }],
  [{ formula: 'B2*2', result: 2401 }],
  [{ value: 'Link', hyperlink: 'https://example.com' }]
]
```

Styles: `font` (`bold` / `color` / `size` / `name`), `fill`, `align`, `valign`, `wrapText`, `numFmt`, `border`.

### PDF

`pdfToJson(input, options?)` → `{ metadata, pageCount, pages, messages }`
`pdfToText(input, options?)` → `string`, pages separated by blank lines
`pdfInfo(input)` → metadata and page count without extracting text (fastest)
`jsonToPdf(spec, options?)` → `Uint8Array`
`pdfPageToPng(input, page?, scale?)` → `Uint8Array` (requires `@napi-rs/canvas`, Node only)

Element types: `text`, `heading`, `image`, `line`, `rect`, `table`, `pageBreak`.
Coordinates use a **top-left origin** (same mental model as HTML): give offsets from the left and top margins.

#### Editing an existing PDF

`editPdf(input, spec)` appends content to an existing PDF and saves it — the original object
structure, content streams, vector art and embedded fonts are kept as-is; only drawing
operators are appended. Also `mergePdfs(inputs)` and `splitPdf(input, { from, to })`.

```javascript
const bytes = await editPdf(file, {
  overlays: [
    [{ type: 'rect', x: 40, y: 60, width: 200, height: 18, fill: '#ffe08a', opacity: 0.5 }],
    [{ type: 'text', x: 40, y: 120, text: 'Approved', size: 14, color: '#334155' }]
  ],
  rotate: [{ page: 0, degrees: 90 }],
  removePages: [4],
  pageOrder: [2, 0, 1]
});
```

`overlays` is indexed by page (0-based) and objects are stacked in array order; `removePages`
and `pageOrder` both act on the final page set (delete first, then reorder). Page management
runs before overlays are drawn, so overlay coordinates always refer to the resulting pages.
Chinese text needs a CJK font via `spec.fonts` — same constraint as `jsonToPdf`.

This is *flattening*, not annotation: overlays are drawn into the page content and cannot be
removed later, which matches the usual "stamp, then export" expectation.

### PPTX

`pptxToHtml`, `pptxToJson`, `pptxToStandard`, `pptxToFiles`, `jsonToPptx`, `PPTXComposer`, `editPptx`

PPTX support comes from `@fefeding/ppt-parser`; this package normalises its input and output types.

## Environment notes

| Scenario | Notes |
| --- | --- |
| Node.js >= 18 | Works out of the box; PDF parsing needs no extra configuration |
| Browser | Requires a bundler (Vite / Webpack / Rollup); for raw ESM CDN use, load `exceljs`'s UMD build instead |
| PDF parsing | `pdfjs-dist` is an optional dependency — only the `pdfTo*` functions need it: `npm i pdfjs-dist` |
| PDF page rendering | Node additionally needs `npm i @napi-rs/canvas` |
| PDFs with Chinese text | Embed a CJK TTF/OTF; the 14 standard PDF fonts contain no CJK glyphs |

```javascript
import fs from 'node:fs';
const bytes = await jsonToPdf({
  fonts: { regular: fs.readFileSync('./SourceHanSans-Regular.ttf') },
  elements: [{ type: 'text', text: '中文内容' }]
});
```

## Known limitations

- **PDF parsing covers text and page info only** — no layout analysis, table structuring or image extraction.
- **DOCX does not restore columns, headers/footers, footnotes, comments or tracked changes**; floating objects (text boxes, SmartArt) are skipped.
- **XLSX does not preserve pivot tables, conditional formatting or comments** (charts and drawings are preserved on round-trip, with their cell-referenced series re-linked on load; comments require `includeNotes`).
- **PPTX fidelity follows `@fefeding/ppt-parser`** (shapes, charts, SmartArt rendering and round-trip).

## Skills

This repository ships agent skills under [`skills/`](./skills) — one package per capability, each with a `SKILL.md` describing when to use it, the exact API, runnable examples and gotchas. They follow the same package model as the work-space skill store (`SKILL.md` + optional resources), so an AI assistant can load them on demand.

| Skill | Covers |
| --- | --- |
| `docx` | parse / generate / round-trip edit DOCX |
| `xlsx` | parse / generate spreadsheets, cells, styles, formulas |
| `pptx` | render / generate / edit presentations |
| `pdf` | parse text, generate, rasterize to PNG, flatten overlays |
| `markdown` | Markdown / plain-text → HTML, Mermaid + KaTeX |
| `convert` | one-call `convert()` between docx / pdf / md / html |
| `assets` | binary resource extraction, per-end access, reference write-back |
| `bundle` | pack / unpack document + assets into a self-contained `.jdbundle` |

## Development

```bash
npm install
npm run test:run      # run tests
npm run type-check    # type check
npm run build         # build ESM / CJS / browser + .d.ts
```

## License

MIT
