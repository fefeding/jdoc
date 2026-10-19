/**
 * jdoc — 文档解析与生成统一库
 *
 * 支持 PPTX / DOCX / XLSX / PDF 四种格式的解析与生成，浏览器与 Node.js 双端可用。
 * 各格式提供独立入口，命名风格统一为：
 *   解析 xxxToYyy(input, options)   生成 jsonToXxx(doc, options)
 *
 * @module jdoc
 */

export * from './core/errors';
export * from './core/input';
export * from './core/xml';
export * from './core/asset';
export * from './core/bundle';

export * from './pptx/index';
export * from './docx/index';
export * from './xlsx/index';
export * from './pdf/index';
export * from './markdown/index';
export * from './text/index';
export * from './convert/index';

/** jdoc 版本号（与 package.json 保持一致） */
export const VERSION = '0.1.0';
