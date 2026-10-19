---
name: jdoc-xlsx
description: 解析、生成与编辑 XLSX 表格（@fefeding/jdoc）。当用户需要把 xlsx 转 JSON/HTML/矩阵、生成电子表格、或处理带样式/公式/超链接的单元格时使用。
---

# XLSX 表格技能（@fefeding/jdoc）

## 何时使用
- 把 `.xlsx` 读成结构化数据（前端表格、图表数据源）
- 用数据生成 Excel 文件
- 处理单元格样式、数字格式、公式、超链接

## 核心 API
| 函数 | 返回 |
| --- | --- |
| `xlsxToJson(input, options?)` | `{ metadata, sheets, messages }` |
| `xlsxToHtml(input, options?)` | 上述 + `html` |
| `xlsxToMatrix(input, sheetName?)` | `{ headers, rows }`，适合喂图表库 |
| `jsonToXlsx(workbook \| sheets, options?)` | `Uint8Array` |
| `sheetsToHtml(workbook \| sheets, options?)` | `string` |

## 单元格三种写法
```js
rows: [
  ['简写：直接写值'],
  [{ 1: '对象式：列号或 A1 记号做键' }, { C2: 'C 列' }],
  [{ value: 1200, style: { numFmt: '#,##0.00', align: 'right' } }],
  [{ formula: 'B2*2', result: 2401 }],
  [{ value: '链接', hyperlink: 'https://example.com' }]
]
```

## 样式
`font`（`bold` / `color` / `size` / `name`）、`fill`、`align`、`valign`、`wrapText`、`numFmt`、`border`。

## 注意事项
- 单元格图片默认 `ignore`；用 `xlsxToJson(file, { assets: 'extract' })` 提取，锚点为 `Sheet1!B3` 这种 A1 记号。
- 不保留数据透视表、图表、条件格式、批注（批注需 `includeNotes`）。
