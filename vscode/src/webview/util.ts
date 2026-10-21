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
export async function resolvePdfWorkerSrc(pdfjs: any): Promise<void> {
    const gwo = pdfjs?.GlobalWorkerOptions;
    if (!gwo) return;
    // 已经是同源 blob（已解析过），无需重复
    if (gwo.workerSrc && String(gwo.workerSrc).startsWith('blob:')) return;
    const raw = (window as any).__JDC_PDF_WORKER__;
    if (!raw) return;
    // 原始地址已就位且并非需要修正的空值：保守不动
    if (gwo.workerSrc === raw) return;
    try {
        const res = await fetch(raw);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const text = await res.text();
        const blob = new Blob([text], { type: 'text/javascript' });
        gwo.workerSrc = URL.createObjectURL(blob);
    } catch {
        // 兜底：某些宿主不允许 fetch 资源时，退回原始地址（沿用旧路径）
        if (!gwo.workerSrc) gwo.workerSrc = raw;
    }
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
