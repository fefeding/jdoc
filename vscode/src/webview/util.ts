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
