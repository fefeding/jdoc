---
name: jdoc-assets
description: 文档内二进制资源（图片等）的统一抽象与取用（@fefeding/jdoc）。当用户需要从解析结果中提取图片字节、按端取合适形式（data URL/Blob/base64/流）、或在生成时按引用写回资源时使用。
---

# 二进制资源 技能（@fefeding/jdoc）

文档里的图片等二进制内容统一抽象为**资源（asset）**：一份字节 + 位置锚点。

## 何时使用
- 解析文档后提取其中的图片字节（存库、转存对象存储、前端显示）
- 在浏览器/Node 不同端取得资源最合适的形式
- 生成文档时按 `assetId` 引用回写，避免搬运 base64

## 资源模型
解析结果中的 `assets` 是 `DocAsset[]`，每项含：
```
{ id, name, mime, data: Uint8Array, part, anchor: { scope, kind, index, path, label } }
```
`anchor` 让「资源 ↔ 文档位置」双向可查。

## 取用助手
| 函数 | 返回 |
| --- | --- |
| `assetToDataUrl(asset)` | `data:image/png;base64,…`（浏览器 `<img src>`） |
| `assetToBlob(asset)` | `Blob`（预览/下载） |
| `assetToBase64(asset)` | base64（存库字段） |
| `assetToArrayBuffer(asset)` | `ArrayBuffer`（底层库/postMessage） |
| `assetToNodeStream(asset)` | Node Readable（直接 pipe 落盘） |
| `assetToWebStream(asset)` | `ReadableStream`（fetch/Response） |

## 解析时提取
```js
const parsed = await docxToJson(file, { assets: 'extract' });
parsed.assets[0].data;   // Uint8Array
```

## 生成时写回（四种格式引用写法一致）
```js
jsonToDocx({ blocks: [{ type: 'image', src: { assetId: 'docx-0001' } }] }, { assets });
jsonToXlsx({ sheets: [{ name: 'S', rows: [['x', { image: { assetId: 'xlsx-0001' } }]] }] }, { assets });
jsonToPptx({ slides: [{ elements: [{ type: 'image', data: { assetId: 'pptx-0001' } }] }] }, { assets });
jsonToPdf({ elements: [{ type: 'image', src: { assetId: 'pdf-0001' } }] }, { assets });
```

## 模式对照
| 模式 | 文档结构里 | `assets` | 适用 |
| --- | --- | --- | --- |
| `inline`（DOCX/XLSX/PPTX 默认） | data URL | 无 | 浏览器直接显示 |
| `extract` | `{ assetId }` 引用 | 有 | 后台存储/另存 |
| `both`（PPTX 默认） | data URL + assetId | 有 | 前后端并行 |
| `ignore`（PDF 默认） | — | 无 | 完全丢弃 |

## 注意事项
- 资源是纯数据（不含方法），可直接 `JSON.stringify` 落库；字节单独取用。
- 同一张图多次出现只存一份字节，块共享同一 `assetId`。
- 资源缺失时 DOCX 降级占位段落、PDF 跳过该图，均不中断生成。
