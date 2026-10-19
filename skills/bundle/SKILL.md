---
name: jdoc-bundle
description: 文档 + 资源 的一体保存与还原（@fefeding/jdoc）。当用户需要把「解析结果 + 二进制资源」打包成自包含文件、或解包还原结构与字节、落盘为可编辑文件夹时使用。
---

# 文档一体保存（Bundle）技能（@fefeding/jdoc）

解析出的文档结构与二进制资源是分离的两份数据。单存 JSON 丢字节、单存字节丢结构——bundle 把两者打成自包含 ZIP，任意机器解压即还原。

## 何时使用
- 把文档「结构 + 资源 + 原始文件」一起持久化（另存为、协作、传输）
- 解包还原成可回传给生成函数的结构
- 在 Node 端落成一个可人工查看/二次加工的文件夹

## 包结构
```
report.jdbundle
├── manifest.json       # 格式、名称、资源清单（含位置锚点）
├── report.docx         # 原始文档（若来源是文件）
└── assets/docx-0001.png
```

## 核心 API
| 函数 | 返回 |
| --- | --- |
| `packDocument(bundle)` | `Uint8Array`（.jdbundle） |
| `unpackDocument(input)` | `{ format, name, document, assets, original? }` |
| `writeDocumentBundle(bundle, dir)` | `string[]`（落盘为文件夹，仅 Node） |
| `isDocumentBundle(input)` | `boolean`（避免误判普通文档） |
| `listBundleAssets(input)` | 只列资源与其在文档中的位置 |

## 典型用法
```js
import { packDocument, unpackDocument, writeDocumentBundle } from '@fefeding/jdoc';

const zip = await packDocument({
  format: 'docx',
  name: '技术协议',
  document: parsed,          // docxToJson 的结果
  assets: parsed.assets,
  original: file             // 可选：保留原始文件
});

const restored = await unpackDocument(zip);   // 还原结构与字节
await writeDocumentBundle(bundle, 'out/技术协议'); // Node 落盘为文件夹
```

## 注意事项
- `DocumentBundle.document` 即各格式解析结果，可直接回传给 `jsonToXxx`。
- 解包后 `assets` 与文档块上的 `assetId` 一一对应，生成时无需再搬运 base64。
- 打包用 JSZip（已是依赖），双端可用；落盘依赖 `node:fs`（仅 Node）。
