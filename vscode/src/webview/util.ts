/** 分块大小：避免把整份文档塞进单次调用（超大数组会爆栈）。 */
const B64_CHUNK = 0x8000;

/** base64 → Uint8Array */
export function b64ToBytes(b64: string): Uint8Array {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    // 分块搬运：与单次 for 循环相比，块边界给 V8 稳定的优化器提示，
    // 100MB 级文档能省下可观的解码时间。
    for (let i = 0; i < bin.length; i += B64_CHUNK) {
        const end = Math.min(i + B64_CHUNK, bin.length);
        for (let j = i; j < end; j++) bytes[j] = bin.charCodeAt(j);
    }
    return bytes;
}

/** Uint8Array → base64 */
export function bytesToB64(bytes: Uint8Array): string {
    let bin = '';
    // 分块：每块用数组收字节再 join，既避开逐字节拼接（O(n) 次 concat 在大文档上极慢），
    // 也避免 String.fromCharCode.apply 把整块一次性展开触碰调用栈参数上限。
    for (let i = 0; i < bytes.length; i += B64_CHUNK) {
        const end = Math.min(i + B64_CHUNK, bytes.length);
        const chars = new Array<string>(end - i);
        for (let j = i, k = 0; j < end; j++, k++) chars[k] = String.fromCharCode(bytes[j]);
        bin += chars.join('');
    }
    return btoa(bin);
}

/**
 * 把 pdfjs 的 worker 地址解析成「同源 Blob URL」后再喂给 pdfjs。
 *
 * 背景：VS Code 官方 webview 的页面 origin（桌面端 vscode-webview://、远端
 * *.vscode-cdn.net）与 asWebviewUri 给出的 worker 地址（桌面端 vscode-resource://、
 * 远端 *.vscode-cdn.net）往往不同源。pdfjs 一旦发现 worker 与页面不同源，会走
 * _createCDNWrapper，用 `await import("<workerSrc>")` 包一层 blob worker —— 跨源
 * import 在 webview 里要么很慢、要么直接失败并回退到主线程 fake worker，表现就是
 * 「打开 PDF 卡很久」。CodeBuddy 由于 webview 与资源同源，跳过包装直接 new Worker，
 * 因此很快。
 *
 * 把 worker 文件 fetch 成文本再包成 Blob URL：blob: 的 origin 与页面同源，pdfjs 判定
 * 同源后跳过 CDN 包装、直接以 module worker 加载，VS Code 官方版与 CodeBuddy 都快。
 * 取到的 worker 是自包含单文件（无相对 import），module worker 可正常执行。
 */
/** 构建脚本内嵌资源时写在 worker 文件开头的标记 */
const EMBEDDED_WORKER_MARKER = '/* JDOC_PDF_EMBEDDED_ASSETS */';
let embeddedPdfAssets = false;
/** worker 是否自带 cMap/字体/wasm（内嵌后不再需要任何资源 URL 可达） */
export function hasEmbeddedPdfAssets(): boolean {
    return embeddedPdfAssets;
}

export async function resolvePdfWorkerSrc(pdfjs: any): Promise<void> {
    const gwo = pdfjs?.GlobalWorkerOptions;
    if (!gwo) return;
    // 已经是同源 blob（已解析过），无需重复
    if (gwo.workerSrc && String(gwo.workerSrc).startsWith('blob:')) return;
    // 宿主注入优先；没注入（旧 HTML / 特殊 fork）就从本脚本地址推导 media/ 下的 worker，
    // 否则 pdfjs 会因缺少 workerSrc 直接失败，页面整片空白。
    let raw = (window as any).__JDC_PDF_WORKER__ as string | undefined;
    if (!raw) {
        for (const s of Array.from(document.querySelectorAll('script[src]'))) {
            const src = (s as HTMLScriptElement).src;
            if (/\/(?:webview|office-editor|pptx-editor)\.js(?:\?|$)/.test(src)) {
                try {
                    raw = new URL('../media/pdf.worker.min.mjs', src).href;
                } catch {
                    // 推导失败就保持 undefined
                }
                break;
            }
        }
    }
    if (!raw) return;
    // 原始地址已就位且并非需要修正的空值：保守不动
    if (gwo.workerSrc === raw) return;
    try {
        const res = await fetch(raw);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const text = await res.text();
        // 新版 worker 开头内嵌了全部 cMap/字体/wasm 并覆盖 fetch；识别到后
        // resolvePdfResourceOptions 会跳过 URL 探测，直接传占位 URL 让拦截器命中。
        embeddedPdfAssets = text.startsWith(EMBEDDED_WORKER_MARKER);
        const blob = new Blob([text], { type: 'text/javascript' });
        gwo.workerSrc = URL.createObjectURL(blob);
    } catch {
        // 兜底：某些宿主不允许 fetch 资源时，退回原始地址（沿用旧路径）
        if (!gwo.workerSrc) gwo.workerSrc = raw;
    }
}

/** 归一化为绝对 URL（worker 内相对路径会抛 Failed to parse URL，必须绝对） */
function absoluteUrl(raw: string): string {
    try {
        return new URL(raw, window.location.href).href;
    } catch {
        return raw;
    }
}

/**
 * media/pdfjs/ 目录的候选基址，按可信度排序。
 *
 * 宿主新版本会注入 __JDC_PDF_CMAPS__ 等；但扩展宿主进程没重启时，webview HTML
 * 仍是旧版本生成的（只有 __JDC_PDF_WORKER__）。该 worker 地址是真实的
 * asWebviewUri（fetch 已验证可达），从它推导同目录下的 pdfjs/ 即可自给自足；
 * 最后再用脚本自身地址兜底。到底哪个可用由 resolvePdfResourceOptions 实测。
 */
function pdfAssetBases(): string[] {
    const w = window as any;
    const bases: string[] = [];
    const push = (u?: string): void => {
        if (!u) return;
        const abs = absoluteUrl(u);
        if (!bases.includes(abs)) bases.push(abs);
    };
    // 1) 宿主注入的 cMap 目录 → 回退到其父目录（.../pdfjs/）
    if (w.__JDC_PDF_CMAPS__) push(String(w.__JDC_PDF_CMAPS__).replace(/cmaps\/?$/, ''));
    // 2) worker 的兄弟目录（旧 HTML 里也有 worker）
    if (w.__JDC_PDF_WORKER__) {
        try {
            push(new URL('pdfjs/', absoluteUrl(String(w.__JDC_PDF_WORKER__))).href);
        } catch {
            // 忽略，继续下一个候选
        }
    }
    // 3) 脚本自身地址推导：dist/webview.js -> ../media/pdfjs/
    for (const s of Array.from(document.querySelectorAll('script[src]'))) {
        const src = (s as HTMLScriptElement).src;
        if (/\/(?:webview|office-editor|pptx-editor)\.js(?:\?|$)/.test(src)) {
            try {
                push(new URL('../media/pdfjs/', src).href);
            } catch {
                break;
            }
        }
    }
    return bases;
}

export interface PdfResourceOptions {
    /** 直接展开进 getDocument 的参数 */
    options: Record<string, unknown>;
    /** 探测结果：cMap 资源确实可 GET 到时为 true */
    reachable: boolean;
    /** 实际采用的 cMap 目录（诊断用） */
    cMapUrl?: string;
}

let cachedResOptions: PdfResourceOptions | null = null;

/**
 * 解析 pdfjs 的 cMap / 标准字体 / wasm 资源目录，并实测可达性。
 *
 * 为什么必须实测：不同宿主（VS Code / CodeBuddy）对 asWebviewUri 的实现与可接受的
 * scheme 不同，推导出的 URL 未必真能 fetch。cMap 探不通时 pdfjs 只打 console
 * warning，页面表现为「文字全白」，所以这里逐个候选探一次，把静默失败变成明确状态。
 *
 * cMap：中文等 CID 字体的 PDF 没有 cMap 时文字整体不渲染（只剩线条/色块）；
 * standard_fonts：未内嵌的 14 种标准字体的替代字形；wasm：JPEG2000 等图像解码。
 */
export async function resolvePdfResourceOptions(): Promise<PdfResourceOptions> {
    if (cachedResOptions) return cachedResOptions;
    const bases = pdfAssetBases();
    const build = (base: string): Record<string, unknown> => ({
        cMapPacked: true,
        useSystemFonts: false,
        cMapUrl: base + 'cmaps/',
        standardFontDataUrl: base + 'standard_fonts/',
        wasmUrl: base + 'wasm/'
    });
    // 内嵌资源的 worker 会拦截自身的 fetch，URL 只是让 pdfjs 拼出文件名的占位符，
    // 不必（也不能跨源）探测；否则 VS Code 官方 webview 里探测失败反而误报。
    if (hasEmbeddedPdfAssets()) {
        const base = bases[0] || absoluteUrl('pdfjs/');
        cachedResOptions = { options: build(base), reachable: true, cMapUrl: base + 'cmaps/' };
        return cachedResOptions;
    }
    for (const base of bases) {
        try {
            // 探一个体积最小的中文 cMap：能 GET 通就说明整套目录可用
            const res = await fetch(base + 'cmaps/UniGB-UCS2-H.bcmap');
            if (res.ok) {
                cachedResOptions = { options: build(base), reachable: true, cMapUrl: base + 'cmaps/' };
                return cachedResOptions;
            }
        } catch {
            // 换下一个候选
        }
    }
    // 全都不通：仍按首个候选传下去（pdfjs 自己还会再试一次），同时标记不可达供 UI 提示
    const fallback = bases[0] ? build(bases[0]) : { cMapPacked: true, useSystemFonts: false };
    cachedResOptions = { options: fallback, reachable: false, cMapUrl: bases[0] ? bases[0] + 'cmaps/' : undefined };
    return cachedResOptions;
}

export function escapeHtml(s: string): string {
    return s.replace(/[&<>"']/g, (c) => {
        switch (c) {
            case '&': return '&amp;';
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '"': return '&quot;';
            default: return '&#39;';
        }
    });
}
