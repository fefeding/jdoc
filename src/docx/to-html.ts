/**
 * DOCX 结构 -> HTML
 *
 * 输出的 HTML 自带内联样式，可直接插入页面预览；调用方也可只取 blocks 自行渲染。
 *
 * @module docx/to-html
 */

import {
    assetToDataUrl,
    getAsset,
    isAssetRef,
    sniffMime,
    toAssetTable,
    type AssetTable,
    type DocAsset
} from '../core/asset';
import type {
    DocxBlock,
    DocxCell,
    DocxCellBorder,
    DocxCellDetail,
    DocxDocument,
    DocxImage,
    DocxList,
    DocxListItem,
    DocxRun,
    DocxTable,
    DocxToc
} from './types';

/** HTML 转义 */
export function escapeHtml(value: string): string {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/** 转义并保留换行（run 内的软换行与制表符） */
function escapeWithBreaks(value: string): string {
    return escapeHtml(value).replace(/\n/g, '<br/>').replace(/\t/g, '&emsp;');
}

/** run -> 内联 span / a */
export function runToHtml(run: DocxRun): string {
    const styles: string[] = [];
    if (run.color) styles.push(`color:${run.color}`);
    if (run.size !== undefined) styles.push(`font-size:${run.size}pt`);
    if (run.font) styles.push(`font-family:${escapeHtml(run.font)}`);
    if (run.highlight) styles.push(`background-color:${run.highlight}`);

    let text = escapeWithBreaks(run.text);
    if (run.bold) text = `<strong>${text}</strong>`;
    if (run.italic) text = `<em>${text}</em>`;
    if (run.underline) text = `<u>${text}</u>`;
    if (run.strike) text = `<s>${text}</s>`;
    if (run.subscript) text = `<sub>${text}</sub>`;
    if (run.superscript) text = `<sup>${text}</sup>`;
    if (run.link) {
        text = `<a href="${escapeHtml(run.link)}" target="_blank" rel="noreferrer noopener">${text}</a>`;
    }
    if (styles.length) return `<span style="${escapeHtml(styles.join(';'))}">${text}</span>`;
    return text;
}

/** runs 或 text -> HTML */
function contentToHtml(text: string | undefined, runs: DocxRun[] | undefined): string {
    if (runs && runs.length) return runs.map(runToHtml).join('');
    return escapeWithBreaks(text || '');
}

/** 列表项 -> HTML（含嵌套子列表） */
function listItemToHtml(item: DocxListItem, ordered: boolean): string {
    const children = typeof item === 'string' ? [] : ((item as { children?: DocxListItem[] }).children || []);
    const self = typeof item === 'string'
        ? escapeWithBreaks(item)
        : (item as DocxRun[]).map(runToHtml).join('');

    if (!children.length) return `<li>${self}</li>`;
    const tag = ordered ? 'ol' : 'ul';
    return `<li>${self}<${tag}>${children.map((child) => listItemToHtml(child, ordered)).join('')}</${tag}></li>`;
}

/** 单元格边框单边 -> CSS 声明 */
function borderSideToCss(side: string, border: DocxCellBorder | undefined, fallback: string): string {
    if (border && border.style === 'nil') return `${side}:none`;
    if (border) {
        const width = border.size !== undefined ? border.size : 1;
        const lineStyle = border.style === 'double' ? 'double' : border.style === 'dashed' ? 'dashed' : border.style === 'dotted' ? 'dotted' : 'solid';
        return `${side}:${width}pt ${lineStyle} ${border.color || '#000000'}`;
    }
    return `${side}:${fallback}`;
}

/** 单元格内容 -> HTML（字符串 / 富文本 / 多段落对象） */
function cellContentToHtml(cell: DocxCell, assets?: AssetTable): string {
    if (typeof cell === 'string') return escapeWithBreaks(cell);
    if (Array.isArray(cell)) return cell.map(runToHtml).join('');
    const detail = cell as DocxCellDetail;
    if (detail.blocks && detail.blocks.length) {
        return detail.blocks.map((block) => blockToHtml(block, assets)).join('');
    }
    if (detail.runs && detail.runs.length) return detail.runs.map(runToHtml).join('');
    return escapeWithBreaks(detail.text || '');
}

/** 单元格 -> td/th（含合并、边框、底纹、对齐、列宽） */
function cellToHtml(cell: DocxCell, isHeader: boolean, defaultBorder: string, assets?: AssetTable): string {
    const tag = isHeader ? 'th' : 'td';
    const content = cellContentToHtml(cell, assets);
    const styles: string[] = [];

    const detail = typeof cell === 'object' && cell !== null && !Array.isArray(cell)
        ? (cell as DocxCellDetail)
        : undefined;

    if (detail) {
        const b = detail.borders;
        styles.push(
            borderSideToCss('border-top', b?.top, defaultBorder),
            borderSideToCss('border-right', b?.right, defaultBorder),
            borderSideToCss('border-bottom', b?.bottom, defaultBorder),
            borderSideToCss('border-left', b?.left, defaultBorder)
        );
        if (detail.fill) styles.push(`background-color:${detail.fill}`);
        if (detail.verticalAlign) styles.push(`vertical-align:${detail.verticalAlign}`);
        if (detail.width) styles.push(`width:${detail.width}pt`);
    } else {
        styles.push(
            borderSideToCss('border-top', undefined, defaultBorder),
            borderSideToCss('border-right', undefined, defaultBorder),
            borderSideToCss('border-bottom', undefined, defaultBorder),
            borderSideToCss('border-left', undefined, defaultBorder)
        );
    }
    styles.push('padding:4px 6px');

    const attrs = [`style="${escapeHtml(styles.join(';'))}"`];
    if (detail?.colspan && detail.colspan > 1) attrs.push(`colspan="${detail.colspan}"`);
    if (detail?.rowspan && detail.rowspan > 1) attrs.push(`rowspan="${detail.rowspan}"`);

    return `<${tag} ${attrs.join(' ')}>${content || '&nbsp;'}</${tag}>`;
}

/** 表格 -> HTML */
function tableToHtml(block: DocxTable, assets?: AssetTable): string {
    const hasHeader = block.header !== false && block.rows.length > 0;
    const columnCount = block.rows.reduce((max, row) => Math.max(max, row.length), 0);
    const defaultBorder = block.borderSize === 0 ? 'none' : '1pt solid #cccccc';

    const colgroup = block.columnWidths
        ? `<colgroup>${block.columnWidths
              .slice(0, columnCount)
              .map((width) => `<col style="width:${width}pt"/>`)
              .join('')}</colgroup>`
        : '';

    const body = block.rows
        .map((row, rowIndex) => {
            const header = hasHeader && rowIndex === 0;
            const cells = row
                .map((cell) => cellToHtml(cell ?? '', header, defaultBorder, assets))
                .join('');
            return `<tr>${cells}</tr>`;
        })
        .join('');

    const tableStyle = block.borderSize === 0
        ? ' style="border-collapse:collapse;border:none"'
        : ' style="border-collapse:collapse"';

    return `<table${tableStyle}>${colgroup}${body}</table>`;
}

/** Word 的 w:numFmt -> CSS list-style-type（未知格式退回 decimal） */
const LIST_STYLE_TYPE: Record<string, string> = {
    decimal: 'decimal',
    decimalzero: 'decimal-leading-zero',
    lowerletter: 'lower-alpha',
    lowerlatin: 'lower-alpha',
    upperletter: 'upper-alpha',
    upperlatin: 'upper-alpha',
    lowerroman: 'lower-roman',
    upperroman: 'upper-roman',
    bullet: 'disc',
    circle: 'circle',
    square: 'square'
};

/** 列表块 -> HTML：编号格式 + 缩进（w:ind 的 left / hanging） */
function listToHtml(block: DocxList): string {
    const tag = block.ordered ? 'ol' : 'ul';
    const styles: string[] = [];
    const fmt = block.numFmt ? LIST_STYLE_TYPE[block.numFmt.toLowerCase()] : undefined;
    // disc 是 <ul> 的初始值，写出来只是冗余噪声，保持标签干净
    if (fmt && fmt !== 'disc') styles.push(`list-style-type:${fmt}`);
    // hanging 是「编号自身占位宽度」：首行左移 hanging，正文仍停在 left 处
    if (block.indentLeft !== undefined) {
        styles.push(`padding-left:${block.indentLeft}pt`);
        if (block.hanging) styles.push(`text-indent:${-block.hanging}pt`);
    }
    const attr = styles.length ? ` style="${escapeHtml(styles.join(';'))}"` : '';
    const items = block.items.map((item) => listItemToHtml(item, block.ordered)).join('');
    return `<${tag}${attr}>${items}</${tag}>`;
}

/** 目录块 -> HTML（层级缩进 + 点线导引 + 右对齐页码） */
function tocToHtml(block: DocxToc): string {
    const rows = block.entries
        .map((entry) => {
            const level = Math.min(Math.max(entry.level, 1), 3);
            const page = entry.page !== undefined ? `<span class="docx-toc-page">${escapeHtml(String(entry.page))}</span>` : '';
            return `<div class="docx-toc-row" data-level="${level}"><span class="docx-toc-text">${escapeHtml(entry.text)}</span><span class="docx-toc-dots"></span>${page}</div>`;
        })
        .join('');
    return `<nav class="docx-toc" aria-label="目录">${rows}</nav>`;
}

/** 段落级属性 -> inline style */
function paragraphStyle(block: DocxBlock): string {
    const para = block as {
        align?: string;
        indentLeft?: number;
        indentRight?: number;
        indentFirstLine?: number;
        spacingBefore?: number;
        spacingAfter?: number;
        lineSpacing?: number;
    };
    const styles: string[] = [];
    if (para.align) styles.push(`text-align:${para.align}`);
    if (para.indentLeft) styles.push(`margin-left:${para.indentLeft}pt`);
    if (para.indentRight) styles.push(`margin-right:${para.indentRight}pt`);
    if (para.indentFirstLine) {
        styles.push(para.indentFirstLine < 0
            ? `text-indent:-${Math.abs(para.indentFirstLine)}pt`
            : `text-indent:${para.indentFirstLine}pt`);
    }
    if (para.spacingBefore) styles.push(`margin-top:${para.spacingBefore}pt`);
    if (para.spacingAfter) styles.push(`margin-bottom:${para.spacingAfter}pt`);
    // Word 的 w:lineRule="auto" 里，w:line 是「单倍行距」的倍数，而单倍行距由字体
    // 行高决定（约 1.2 × 字号），并非字号本身。CSS 的无单位 line-height 却是相对
    // font-size 的，直接照搬会让行距偏小约 20%（中文文档尤明显）。
    if (para.lineSpacing && para.lineSpacing !== 1) {
        styles.push(`line-height:${Math.round(para.lineSpacing * 1.2 * 10000) / 10000}`);
    }
    return styles.length ? ` style="${escapeHtml(styles.join(';'))}"` : '';
}

/** 单块 -> HTML */
/** 图片 src -> 可放进 <img src> 的地址：内联字符串直用，资源引用回查资源表转 data URL */
function resolveImageHtmlSrc(src: DocxImage['src'], assets?: AssetTable): string | undefined {
    if (typeof src === 'string') return src;
    if (isAssetRef(src)) {
        const ref = src as { assetId?: string; asset?: DocAsset };
        const asset = ref.asset ?? getAsset(assets, ref.assetId);
        return asset ? assetToDataUrl(asset) : undefined;
    }
    if (src instanceof Uint8Array || src instanceof ArrayBuffer) {
        const bytes = src instanceof Uint8Array ? src : new Uint8Array(src);
        return assetToDataUrl(bytes, sniffMime(bytes));
    }
    return undefined;
}

export function blockToHtml(block: DocxBlock, assets?: AssetTable): string {
    switch (block.type) {
        case 'heading': {
            const level = block.level >= 1 && block.level <= 6 ? block.level : 3;
            // 章节号由 numbering 生成、不在文本里，这里前置还原（标题与目录条目才对得上）
            const number = block.number
                ? `<span class="docx-h-num">${escapeHtml(block.number)}</span>`
                : '';
            return `<h${level}${paragraphStyle(block)}>${number}${contentToHtml(block.text, block.runs)}</h${level}>`;
        }
        case 'paragraph':
            return `<p${paragraphStyle(block)}>${contentToHtml(block.text, block.runs) || '&nbsp;'}</p>`;
        case 'quote':
            return `<blockquote${paragraphStyle(block)}>${contentToHtml(block.text, block.runs)}</blockquote>`;
        case 'code':
            return `<pre style="background:#f5f5f5;padding:8px;border-radius:4px;overflow:auto"><code>${escapeHtml(block.text)}</code></pre>`;
        case 'list':
            return listToHtml(block);
        case 'table':
            return tableToHtml(block, assets);
        case 'image': {
            // 资源引用（{ assetId }）需回查资源表取字节，转成 data URL 才能在 <img> 里显示
            const resolved = resolveImageHtmlSrc(block.src, assets);
            if (!resolved) return '';
            const attrs = [
                block.width ? `width="${Math.round(block.width * (96 / 72))}"` : '',
                block.height ? `height="${Math.round(block.height * (96 / 72))}"` : '',
                block.alt ? `alt="${escapeHtml(block.alt)}"` : 'alt=""'
            ]
                .filter(Boolean)
                .join(' ');
            const img = `<img src="${escapeHtml(resolved)}" ${attrs} style="max-width:100%"/>`;
            const caption = block.caption
                ? `<figcaption style="text-align:center;font-size:11px;color:#666">${escapeHtml(block.caption)}</figcaption>`
                : '';
            return `<figure style="margin:12px 0;text-align:center">${img}${caption}</figure>`;
        }
        case 'textBox': {
            const styles: string[] = ['margin:8px 0', 'padding:8px 12px', 'overflow:auto'];
            if (block.fill) {
                styles.push(`background-color:${block.fill}`);
                // 深色底自动反白（run 未显式指定颜色时保证可读）
                const hex = block.fill.replace(/^#/, '');
                if (/^[0-9a-fA-F]{6}$/.test(hex)) {
                    const r = parseInt(hex.slice(0, 2), 16);
                    const g = parseInt(hex.slice(2, 4), 16);
                    const b = parseInt(hex.slice(4, 6), 16);
                    if (r * 0.299 + g * 0.587 + b * 0.114 < 128) styles.push('color:#ffffff');
                }
            }
            if (block.stroke) styles.push(`border:1px solid ${block.stroke}`);
            if (block.width) styles.push(`width:${block.width}pt`);
            const inner = block.paragraphs.map((item) => blockToHtml(item, assets)).join('');
            return `<figure class="docx-textbox" style="${escapeHtml(styles.join(';'))}">${inner}</figure>`;
        }
        case 'pageBreak':
            return '<hr style="page-break-after:always;border:none;border-top:1px dashed #ccc"/>';
        case 'hr':
            return '<hr/>';
        case 'toc':
            return tocToHtml(block);
        default:
            return '';
    }
}

/** 默认样式表（便于开箱即用，也可忽略自带样式自行渲染） */
export const DOCX_HTML_STYLE = `.docx-body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,"PingFang SC","Microsoft YaHei",sans-serif;font-size:11pt;line-height:1.6;color:#1f2328}
.docx-page{background:#fff;margin:24px auto;box-shadow:0 1px 6px rgba(0,0,0,.2);box-sizing:border-box}
.docx-body h1,.docx-body h2,.docx-body h3,.docx-body h4,.docx-body h5,.docx-body h6{font-weight:700;line-height:1.35;color:#1f2328;break-after:avoid}
.docx-body h1{font-size:20pt;margin:22px 0 12px}
.docx-body h2{font-size:16pt;margin:20px 0 10px}
.docx-body h3{font-size:14pt;margin:18px 0 9px}
.docx-body h4{font-size:12.5pt;margin:16px 0 8px}
.docx-body h5,.docx-body h6{font-size:11.5pt;margin:14px 0 7px;font-weight:600}
.docx-h-num{margin-right:.45em;white-space:nowrap}
.docx-body p{margin:0}
.docx-body blockquote{margin:8px 0;padding:4px 12px;border-left:3px solid #d0d7de;color:#57606a}
.docx-body table{margin:10px 0;font-size:10.5pt}
.docx-body th{background:#f6f8fa;font-weight:600}
.docx-body img{max-width:100%}
.docx-textbox{margin:8px 0;font-size:10pt;line-height:1.45}
.docx-textbox p{margin:0}
.docx-toc{margin:6px 0 14px}
.docx-toc-row{display:flex;align-items:baseline;gap:.3em;padding:2px 0;line-height:1.5}
.docx-toc-row[data-level="1"]{font-weight:600;margin-top:4px}
.docx-toc-row[data-level="2"]{padding-left:1.6em}
.docx-toc-row[data-level="3"]{padding-left:3.2em;font-size:.95em}
.docx-toc-text{white-space:pre}
.docx-toc-dots{flex:1 1 auto;border-bottom:1px dotted currentColor;opacity:.45;transform:translateY(-.28em);min-width:1.5em}
.docx-toc-page{white-space:nowrap}
`;

/**
 * 把 DOCX 结构渲染为 HTML
 * @param input - DocxDocument 或 blocks 数组（传 DocxDocument 时按 section 还原页面）
 * @param options.includeStyle - 是否附带默认样式表（默认 true）
 */
export function docxToHtmlString(
    input: DocxDocument | DocxBlock[],
    options: { includeStyle?: boolean } = {}
): string {
    const doc = Array.isArray(input) ? undefined : input;
    const blocks = Array.isArray(input) ? input : input?.blocks || [];
    const assets = doc?.assets?.length ? toAssetTable(doc.assets) : undefined;
    const body = blocks.map((block) => blockToHtml(block, assets)).filter(Boolean).join('\n');

    // 有页面设置时还原纸张：固定页宽 + 页边距，接近 Word 的版面
    let pageOpen = '<div class="docx-body">';
    let pageClose = '</div>';
    const section = doc?.section;
    if (section) {
        pageOpen =
            `<div class="docx-page" style="width:${section.width}pt;min-height:${section.height}pt;` +
            `padding:${section.marginTop}pt ${section.marginRight}pt ${section.marginBottom}pt ${section.marginLeft}pt">` +
            '<div class="docx-body">';
        pageClose = '</div></div>';
    }

    // 中文 Word/WPS 几乎都会开「行网格」（w:docGrid w:type="lines"），未显式设行距的
    // 段落其单倍行距就是 linePitch。不还原它会用 CSS 的 1.6 倍，行距偏小、每页塞进
    // 更多内容，分页落点整体比 Word/WPS 前移。这里只改默认值，显式带 lineSpacing 的
    // 段落仍由 paragraphStyle 的内联样式覆盖。
    // 写成 CSS 规则而非 .docx-body 的内联样式：分页场景会把块搬到新建的 .docx-body
    // 上，内联样式随之丢失，而随 HTML 返回的 <style> 会被提取后长期生效。
    const gridCss = section?.gridLinePitch ? `.docx-body{line-height:${section.gridLinePitch}pt}` : '';
    const style = options.includeStyle === false ? '' : `<style>${DOCX_HTML_STYLE}${gridCss}</style>`;
    return `${style}${pageOpen}${body}${pageClose}`;
}
