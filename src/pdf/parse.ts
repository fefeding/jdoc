/**
 * PDF 解析
 *
 * 基于 pdfjs-dist 提取纯文本、页面尺寸与文本坐标。
 * pdfjs-dist 是 optionalDependencies（体积较大且含 worker），因此这里用动态导入，
 * 未安装时给出明确的安装指引，而不是抛一个费解的模块解析错误。
 *
 * @module pdf/parse
 */

import { JdocError, wrapError } from '../core/errors';
import { toUint8Array, type DocInput } from '../core/input';
import type { DocAsset } from '../core/asset';
import { extractPdfAssets } from './assets';
import type { PdfDocument, PdfMetadata, PdfPage, PdfParseOptions, PdfTextItem } from './types';

/** pdfjs 的文本项结构（避免直接依赖 pdfjs-dist 的类型声明） */
interface PdfjsTextItem {
    str: string;
    dir: string;
    width: number;
    height: number;
    transform: number[];
    fontName: string;
    hasEOL: boolean;
}

interface PdfjsPage {
    getViewport(options: { scale: number }): { width: number; height: number };
    getTextContent(): Promise<{ items: (PdfjsTextItem | unknown)[] }>;
    rotate: number;
    getOperatorList?(): Promise<{ fnArray: number | string[]; argsArray?: unknown[][] }>;
    objs?: { get(name: string): unknown };
    commonObjs?: { get(name: string): unknown };
}

interface PdfjsDocument {
    numPages: number;
    getPage(pageNumber: number): Promise<PdfjsPage>;
    getAttachments?(): Promise<unknown>;
    getAttachmentContent?(filename: string): Promise<Uint8Array>;
    getMetadata(): Promise<{ info: Record<string, unknown> }>;
    cleanup?(): Promise<void>;
    destroy?(): Promise<void>;
}

interface PdfjsModule {
    getDocument(params: Record<string, unknown>): PdfjsLoadingTask;
    GlobalWorkerOptions?: { workerSrc: string };
}

/** pdfjs 的加载任务句柄 */
interface PdfjsLoadingTask {
    promise: Promise<PdfjsDocument>;
    destroy?: () => void;
}

/** 尝试加载 pdfjs-dist，失败时抛出带安装指引的错误 */
async function loadPdfjs(): Promise<PdfjsModule> {
    const errors: string[] = [];
    const attempt = async (loader: () => Promise<unknown>, label: string) => {
        try {
            return (await loader()) as PdfjsModule;
        } catch (error) {
            errors.push(`${label}: ${(error as Error)?.message || error}`);
            return null;
        }
    };
    // 注意：必须使用字面量说明符（而非变量），否则打包器（Vite/Rollup）无法在浏览器
    // 构建期将 pdfjs 打包进产物，运行时裸说明符会被浏览器拒绝。
    const mod =
        (await attempt(() => import('pdfjs-dist/legacy/build/pdf.mjs'), 'pdfjs-dist/legacy/build/pdf.mjs')) ||
        (await attempt(() => import('pdfjs-dist/build/pdf.mjs'), 'pdfjs-dist/build/pdf.mjs')) ||
        (await attempt(() => import('pdfjs-dist'), 'pdfjs-dist'));
    if (mod) return mod;
    throw new JdocError(
        'PDF 解析需要 pdfjs-dist，请先安装：npm i pdfjs-dist（或 npm i @fefeding/jdoc）',
        'pdf',
        { cause: errors.join(' | ') }
    );
}

/** PDF 日期串（形如 D:20240101120000+08'00'）-> ISO 字符串 */
function parsePdfDate(value: unknown): string | undefined {
    if (!value) return undefined;
    const raw = String(value);
    const match = /D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(raw);
    if (!match) return undefined;

    const [, year, month = '01', day = '01', hour = '00', minute = '00', second = '00'] = match;
    // 尾部时区：+08'00' / Z / -05'00'
    const tzMatch = /([+-Z])(\d{2})?'?(\d{2})?'?/.exec(raw.slice(raw.indexOf(match[0]) + match[0].length));
    let iso = `${year}-${month}-${day}T${hour}:${minute}:${second}`;
    if (tzMatch && tzMatch[1] !== 'Z') {
        iso += `${tzMatch[1]}${tzMatch[2] || '00'}:${tzMatch[3] || '00'}`;
    } else if (tzMatch && tzMatch[1] === 'Z') {
        iso += 'Z';
    } else {
        iso += 'Z';
    }
    const date = new Date(iso);
    return Number.isNaN(date.getTime()) ? undefined : date.toISOString();
}

/** 解析 PDF 文档信息字典为统一元数据 */
function readMetadata(info: Record<string, unknown>): PdfMetadata | undefined {
    const metadata: PdfMetadata = {};
    const put = (key: keyof PdfMetadata, value: unknown) => {
        if (value === undefined || value === null || value === '') return;
        (metadata as Record<string, string>)[key] = String(value);
    };

    put('title', info.Title);
    put('author', info.Author);
    put('subject', info.Subject);
    put('keywords', info.Keywords);
    put('creator', info.Creator);
    put('producer', info.Producer);
    put('version', info.PDFFormatVersion);

    const created = parsePdfDate(info.CreationDate);
    const modified = parsePdfDate(info.ModDate);
    if (created) metadata.created = created;
    if (modified) metadata.modified = modified;

    return Object.keys(metadata).length ? metadata : undefined;
}

/**
 * 探测 pdfjs-dist 自带的标准字体 / cMap 目录（Node 下自动生效，浏览器需手动传入）。
 *
 * 刻意用变量承接 specifier 再动态 import：这样打包器不会在浏览器构建期
 * 尝试解析 node: 内置模块并报 externalize 警告。
 */
async function resolveResourceUrl(subDir: string): Promise<string | undefined> {
    const isNode = typeof process !== 'undefined' && !!process.versions?.node;
    if (!isNode) return undefined;

    const moduleSpecifier = 'node:module';
    const urlSpecifier = 'node:url';
    try {
        const { createRequire } = (await import(/* @vite-ignore */ moduleSpecifier)) as {
            createRequire: (from: string) => { resolve: (id: string) => string };
        };
        const { pathToFileURL } = (await import(/* @vite-ignore */ urlSpecifier)) as {
            pathToFileURL: (path: string) => URL;
        };

        // 从当前工作目录解析，命中项目里的 pdfjs-dist
        const require = createRequire(pathToFileURL(process.cwd() + '/').href);
        const resolved = require.resolve('pdfjs-dist/package.json');
        return `${resolved.replace(/package\.json$/, '')}${subDir}/`;
    } catch {
        // 探测失败不影响主流程：仅标准字体解码与中文 cMap 会退化
        return undefined;
    }
}

/** 把 pdfjs 文本项转换为统一结构 */
function toTextItem(item: PdfjsTextItem): PdfTextItem {
    const transform = item.transform || [];
    return {
        text: item.str,
        // transform = [a, b, c, d, e, f]，其中 e/f 为基线起点的 x/y
        x: transform[4] || 0,
        y: transform[5] || 0,
        width: item.width || 0,
        height: item.height || 0,
        // 变换矩阵的纵向缩放即字号
        fontSize: Math.abs(transform[3] || item.height || 0),
        fontName: item.fontName,
        hasEOL: item.hasEOL
    };
}

/** 文本项数组 -> 整页纯文本（按 hasEOL 与坐标变化推断换行） */
function itemsToText(items: PdfTextItem[]): string {
    let text = '';
    let lastBottom: number | null = null;

    for (const item of items) {
        // 行距显著变化说明换行；否则补一个空格避免中文/单词粘连
        const gap = lastBottom === null ? 0 : item.y - lastBottom;
        if (lastBottom !== null && (item.hasEOL || Math.abs(gap) > item.fontSize * 0.5)) {
            text += '\n';
        } else if (lastBottom !== null && /[一-龥]$/.test(text) === false && /^[一-龥]/.test(item.text)) {
            text += ' ';
        }
        text += item.text;
        lastBottom = item.y;
        if (item.hasEOL) text += '\n';
    }

    return text.replace(/\n{3,}/g, '\n\n').trim();
}

/** 2D 仿射矩阵相乘（a·b），输入/输出均为 [a,b,c,d,e,f] */
function mulMatrix(a: number[], b: number[]): number[] {
    return [
        a[0] * b[0] + a[2] * b[1],
        a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3],
        a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4],
        a[1] * b[4] + a[3] * b[5] + a[5]
    ];
}

/**
 * 从页面的 operator list 提取矢量线段（表格线 / 分隔线 / 矩形边框）
 *
 * 逐 op 维护当前变换矩阵（CTM），把 moveTo/lineTo/rectangle/constructPath 产生的
 * 线段转成 PDF 用户坐标下的线段；文本与图像 op 不影响路径，直接跳过。
 * 解析失败（无 operator list 或被加密保护）返回空数组。
 */
async function extractVectors(
    pdfjs: PdfjsModule,
    page: PdfjsPage
): Promise<import('./types').PdfVector[]> {
    const opList = (await page.getOperatorList?.()) as unknown as
        | { fnArray: number[]; argsArray: unknown[][] }
        | undefined;
    if (!opList) return [];

    const OPS = (pdfjs as unknown as { OPS?: Record<string, number> }).OPS;
    if (!OPS) return [];

    const lines: import('./types').PdfVector[] = [];
    let ctm = [1, 0, 0, 1, 0, 0];
    const stack: number[][] = [];
    let sub: number[][] = [];
    let closed = false;

    const transform = (x: number, y: number): [number, number] => [
        ctm[0] * x + ctm[2] * y + ctm[4],
        ctm[1] * x + ctm[3] * y + ctm[5]
    ];
    const pushPoint = (x: number, y: number) => {
        sub.push(transform(x, y));
    };
    const flush = () => {
        for (let i = 1; i < sub.length; i++) addLine(lines, sub[i - 1], sub[i]);
        if (closed && sub.length > 1) addLine(lines, sub[sub.length - 1], sub[0]);
        sub = [];
        closed = false;
    };
    const addRectPath = (x: number, y: number, w: number, h: number) => {
        sub = [];
        pushPoint(x, y);
        pushPoint(x + w, y);
        pushPoint(x + w, y + h);
        pushPoint(x, y + h);
        closed = true;
    };

    const args = (opList.argsArray || []) as unknown[][];
    for (let i = 0; i < opList.fnArray.length; i++) {
        const op = opList.fnArray[i];
        const a = (args[i] as unknown[]) || [];
        if (op === OPS.save) stack.push(ctm.slice());
        else if (op === OPS.restore) ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
        else if (op === OPS.transform) ctm = mulMatrix(a as unknown as number[], ctm);
        else if (op === OPS.paintFormXObjectBegin) {
            stack.push(ctm.slice());
            ctm = mulMatrix((a[0] as unknown as number[]) || [1, 0, 0, 1, 0, 0], ctm);
        } else if (op === OPS.paintFormXObjectEnd) ctm = stack.pop() || ctm;
        else if (op === OPS.rectangle) addRectPath(a[0] as number, a[1] as number, a[2] as number, a[3] as number);
        else if (op === OPS.constructPath) {
            const ops = a[0] as unknown as number[];
            const coords = a[1] as unknown as number[];
            let p = 0;
            sub = [];
            for (const code of ops) {
                if (code === 0) {
                    pushPoint(coords[p], coords[p + 1]);
                    p += 2;
                } else if (code === 1) {
                    pushPoint(coords[p], coords[p + 1]);
                    p += 2;
                } else if (code === 2) {
                    p += 6; // 曲线：用终点近似（表格还原不需要曲线精度）
                } else if (code === 3) {
                    p += 4;
                } else if (code === 4) {
                    closed = true;
                }
            }
        } else if (
            op === OPS.stroke ||
            op === OPS.fill ||
            op === OPS.closeStroke ||
            op === OPS.closeFillStroke ||
            op === OPS.eoFill ||
            op === OPS.eoFillStroke
        ) {
            flush();
        } else if (op === OPS.endPath) {
            sub = [];
            closed = false;
        }
    }
    return lines;
}

/** 添加一条线段（过滤退化点），标注水平/垂直 */
function addLine(lines: import('./types').PdfVector[], p1: number[], p2: number[]): void {
    const [x1, y1] = p1;
    const [x2, y2] = p2;
    if (Math.hypot(x2 - x1, y2 - y1) < 0.5) return;
    const dx = Math.abs(x2 - x1);
    const dy = Math.abs(y2 - y1);
    lines.push({
        x1,
        y1,
        x2,
        y2,
        horizontal: dy <= 0.6 && dx > 1,
        vertical: dx <= 0.6 && dy > 1
    });
}

/**
 * 解析 PDF
 * @param input - ArrayBuffer / Uint8Array / Buffer / Blob / base64 字符串
 * @param options - 解析选项
 */
export async function parsePdf(
    input: DocInput,
    options: PdfParseOptions = {}
): Promise<{ document: PdfDocument; messages: string[] }> {
    const messages: string[] = [];
    let pdf: PdfjsDocument | undefined;
    let loadingTask: PdfjsLoadingTask | undefined;

    try {
        const bytes = await toUint8Array(input);
        if (bytes.byteLength < 5) throw new TypeError('文件过小');
        // %PDF- 魔数校验，避免把损坏文件交给 pdfjs
        if (bytes[0] !== 0x25 || bytes[1] !== 0x50 || bytes[2] !== 0x44 || bytes[3] !== 0x46) {
            throw new TypeError('不是有效的 PDF 文件（缺少 %PDF- 头）');
        }

        const pdfjs = await loadPdfjs();
        // 浏览器端 pdfjs 必须配置 worker。打包器（Vite/Rollup）会把该 URL 表达式
        // 解析为构建产物中的 worker 资源地址；Node 端无需 worker，直接跳过。
        if (typeof window !== 'undefined' && pdfjs.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
            try {
                pdfjs.GlobalWorkerOptions.workerSrc = new URL(
                    'pdfjs-dist/legacy/build/pdf.worker.min.mjs',
                    import.meta.url
                ).toString();
            } catch {
                // 资源解析失败不阻断：调用方可自行设置 GlobalWorkerOptions.workerSrc
            }
        }
        const standardFontDataUrl = options.standardFontDataUrl || (await resolveResourceUrl('standard_fonts'));
        const cMapUrl = await resolveResourceUrl('cmaps');

        loadingTask = pdfjs.getDocument({
            // 必须传副本：pdfjs 会把 data 的 ArrayBuffer transfer 给 worker，
            // 原始 bytes 会随之 detach（长度归零），调用方手里的字节就废了。
            // 同一份字节再解析一次会直接命中「文件过小」。
            data: bytes.slice(),
            standardFontDataUrl,
            cMapUrl,
            // Node 与受限浏览器环境下禁用字体探测与 JIT，减少兼容问题
            isEvalSupported: false,
            useSystemFonts: false
        });
        pdf = await loadingTask.promise;

        const metadata = readMetadata((await pdf.getMetadata()).info || {});
        const document: PdfDocument = {
            metadata,
            pageCount: pdf.numPages
        };

        // 内嵌资源（附件 / 页面图像）：默认关闭，开启后按资源表返回
        if (options.assets && options.assets !== 'ignore') {
            try {
                const assets: DocAsset[] = await extractPdfAssets(pdf as never, {
                    mode: options.assets,
                    maxPages: options.maxPages,
                    pages: options.pages
                });
                if (assets.length) document.assets = assets;
            } catch (error) {
                messages.push(`内嵌资源提取失败：${(error as Error)?.message || error}`);
            }
        }

        // 仅在需要页面内容时才逐页加载，避免大文件无谓开销
        const wantsContent = options.text !== false || options.items === true;
        if (wantsContent) {
            const wanted = new Set(options.pages || []);
            const limit = options.maxPages || pdf.numPages;
            const pages: PdfPage[] = [];

            for (let index = 1; index <= pdf.numPages && pages.length < limit; index++) {
                if (wanted.size && !wanted.has(index)) continue;

                const page = await pdf.getPage(index);
                const viewport = page.getViewport({ scale: 1 });
                const record: PdfPage = {
                    index,
                    width: Math.round(viewport.width * 100) / 100,
                    height: Math.round(viewport.height * 100) / 100,
                    rotation: ((page.rotate % 360) + 360) % 360
                };

                if (options.text !== false || options.items) {
                    const content = await page.getTextContent();
                    const items = (content.items as (PdfjsTextItem | unknown)[])
                        .filter((item): item is PdfjsTextItem => !!item && typeof (item as PdfjsTextItem).str === 'string')
                        .map(toTextItem);

                    if (options.text !== false) record.text = itemsToText(items);
                    if (options.items) record.items = items;
                }

                if (options.vectors) {
                    try {
                        record.vectors = await extractVectors(pdfjs, page);
                    } catch {
                        // 矢量提取失败不影响文本主流程
                    }
                }

                pages.push(record);
            }

            if (wanted.size && pages.length < wanted.size) {
                messages.push(`部分页码超出文档范围（共 ${pdf.numPages} 页）`);
            }
            document.pages = pages;
        }

        return { document, messages };
    } catch (error) {
        throw wrapError('pdf', '解析', error);
    } finally {
        // pdfjs 6.x 中 destroy 挂在 loadingTask 上，cleanup 挂在 document 上
        try {
            if (pdf?.cleanup) await pdf.cleanup();
            else if (pdf?.destroy) await pdf.destroy();
            else if (loadingTask?.destroy) loadingTask.destroy();
        } catch {
            // 清理失败不影响结果
        }
    }
}
