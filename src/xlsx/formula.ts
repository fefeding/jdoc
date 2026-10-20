/**
 * XLSX 公式引擎
 *
 * 把单元格公式（不含前导 `=`）求值成结果，供编辑器增量重算与图表取数共用。
 *
 * 覆盖范围：
 * - 运算与比较：`+ - * / ^ % &`、`< <= > >= = <>`
 * - 引用：单元格（`B3`）、区域（`B3:C9`）、跨表（`销售数据!G3:G32`、`'销售 表'!A1`）
 * - 函数：统计（SUM/AVERAGE/COUNT/COUNTA/COUNTIF/SUMIF/…）、逻辑（IF/AND/OR/IFERROR/…）、
 *   文本（LEFT/RIGHT/MID/LEN/TEXT/CONCAT/…）、查找（VLOOKUP/INDEX/MATCH）、
 *   日期（DATE/YEAR/MONTH/DAY/TEXT）、数学（ABS/ROUND/SQRT/MOD/…）
 *
 * 关键约定：
 * - 表名允许任意字符（中文、空格），靠「标识符后紧跟 `!`」判定表名，
 *   因此中文表名不会被误当成普通标识符丢掉（这是旧实现最大的坑）。
 * - 区域引用在函数实参里保留为二维结构，供 SUMIF / VLOOKUP 等按位置取数。
 * - 未知函数、未知表名会抛错，由调用方决定是否保留原缓存值，
 *   避免「算不出来就把正确结果覆盖成 0」。
 *
 * @module xlsx/formula
 */

import { formatByPattern } from '../core/numfmt';

/** 公式可用的标量值 */
export type FormulaValue = number | string | boolean;

/** 单元格读取结果 */
export interface FormulaCellInfo {
    /** 单元格值（公式单元格取计算结果） */
    value?: FormulaValue | null;
    /** 显示文本 */
    text?: string;
}

/** 求值上下文 */
export interface FormulaContext {
    /** 当前所在表名（相对引用的基准） */
    sheetName: string;
    /** 读取单元格；row/col 均 0 基 */
    read(sheetName: string, row: number, col: number): FormulaCellInfo;
    /** 判断表是否存在；缺省视为都存在 */
    hasSheet?(sheetName: string): boolean;
}

/** 公式求值失败（调用方可据此保留原值） */
export class FormulaError extends Error {
    constructor(public readonly code: string) {
        super(code);
        this.name = 'FormulaError';
    }
}

// ---------- A1 记号 ----------

/** 列号（1 基）-> 列字母 */
export function columnName(index: number): string {
    let n = Math.max(1, Math.floor(index));
    let letters = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        letters = String.fromCharCode(65 + rem) + letters;
        n = Math.floor((n - 1) / 26);
    }
    return letters;
}

/** 列字母 -> 列号（1 基） */
export function columnIndex(letters: string): number {
    let index = 0;
    for (const ch of String(letters).toUpperCase()) {
        if (ch < 'A' || ch > 'Z') return NaN;
        index = index * 26 + (ch.charCodeAt(0) - 64);
    }
    return index;
}

/** 'B3' -> { row: 2, col: 1 }（0 基） */
export function parseCellRef(ref: string): { row: number; col: number } | null {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref).trim());
    if (!m) return null;
    const col = columnIndex(m[1]);
    const row = parseInt(m[2], 10);
    if (!Number.isFinite(col) || !Number.isFinite(row) || row < 1) return null;
    return { row: row - 1, col: col - 1 };
}

/**
 * 解析区域引用（可带表名）
 *
 * 支持 `'汇总看板'!$B$8:$B$10`、`销售数据!G3:G32`、`B3:C9` 三种写法。
 */
export function parseRangeRef(ref: string): {
    sheet?: string;
    r1: number;
    c1: number;
    r2: number;
    c2: number;
} | null {
    const text = String(ref).trim();
    let sheet: string | undefined;
    let body = text;

    const bang = body.lastIndexOf('!');
    if (bang >= 0) {
        const head = body.slice(0, bang).trim();
        body = body.slice(bang + 1).trim();
        if (head.startsWith("'") && head.endsWith("'") && head.length >= 2) {
            sheet = head.slice(1, -1).replace(/''/g, "'");
        } else {
            sheet = head;
        }
    }

    const parts = body.split(':');
    const a = parseCellRef(parts[0]);
    if (!a) return null;
    const b = parts[1] ? parseCellRef(parts[1]) : a;
    if (!b) return null;
    return {
        sheet,
        r1: Math.min(a.row, b.row),
        c1: Math.min(a.col, b.col),
        r2: Math.max(a.row, b.row),
        c2: Math.max(a.col, b.col)
    };
}

// ---------- 词法 ----------

type Token =
    | { kind: 'num'; value: number }
    | { kind: 'str'; value: string }
    | { kind: 'bool'; value: boolean }
    | { kind: 'ref'; sheet?: string; ref: string }
    | { kind: 'range'; sheet?: string; from: string; to: string }
    | { kind: 'func'; name: string }
    | { kind: 'op'; op: string }
    | { kind: 'name'; name: string };

const IDENT_START = /[\p{L}_\\]/u;
const IDENT_PART = /[\p{L}\p{N}_\\.]/u;
const CELL_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/;

function isSpace(ch: string): boolean {
    return ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r';
}

/** 读取一个 A1 引用（可带 $），返回 [ref, 结束位置] */
function readRefToken(src: string, start: number): { ref: string; end: number } | null {
    const m = /^\$?[A-Za-z]{1,3}\$?\d+/.exec(src.slice(start));
    if (!m) return null;
    return { ref: m[0].replace(/\$/g, ''), end: start + m[0].length };
}

/** 词法分析 */
export function tokenizeFormula(src: string): Token[] {
    const text = String(src).replace(/^=/, '');
    const toks: Token[] = [];
    let i = 0;

    while (i < text.length) {
        const ch = text[i];

        if (isSpace(ch)) {
            i++;
            continue;
        }

        // 字符串："..."（"" 表示一个引号）
        if (ch === '"') {
            let out = '';
            i++;
            while (i < text.length) {
                if (text[i] === '"') {
                    if (text[i + 1] === '"') {
                        out += '"';
                        i += 2;
                        continue;
                    }
                    i++;
                    break;
                }
                out += text[i++];
            }
            toks.push({ kind: 'str', value: out });
            continue;
        }

        // 带引号的表名：'销售 数据'!A1
        if (ch === "'") {
            let out = '';
            i++;
            while (i < text.length) {
                if (text[i] === "'") {
                    if (text[i + 1] === "'") {
                        out += "'";
                        i += 2;
                        continue;
                    }
                    i++;
                    break;
                }
                out += text[i++];
            }
            if (text[i] === '!') {
                i++;
                const ref = readRefToken(text, i);
                if (!ref) throw new FormulaError('#REF!');
                i = ref.end;
                const range = readRangeTail(text, i, out, ref.ref);
                toks.push(range.token);
                i = range.end;
            } else {
                toks.push({ kind: 'str', value: out });
            }
            continue;
        }

        // 数字
        if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(text[i + 1] || ''))) {
            const m = /^\d*\.?\d+(?:[eE][+-]?\d+)?/.exec(text.slice(i))!;
            toks.push({ kind: 'num', value: parseFloat(m[0]) });
            i += m[0].length;
            continue;
        }

        // 标识符：函数名 / 表名 / 单元格引用 / 定义名称
        if (IDENT_START.test(ch)) {
            let j = i + 1;
            while (j < text.length && IDENT_PART.test(text[j])) j++;
            const word = text.slice(i, j);
            const upper = word.toUpperCase();

            // 函数调用
            let k = j;
            while (k < text.length && isSpace(text[k])) k++;
            if (text[k] === '(') {
                toks.push({ kind: 'func', name: upper });
                i = j;
                continue;
            }

            if (upper === 'TRUE' || upper === 'FALSE') {
                toks.push({ kind: 'bool', value: upper === 'TRUE' });
                i = j;
                continue;
            }

            // 跨表引用：标识符后紧跟 '!'
            if (text[k] === '!') {
                i = k + 1;
                const ref = readRefToken(text, i);
                if (!ref) throw new FormulaError('#REF!');
                i = ref.end;
                const range = readRangeTail(text, i, word, ref.ref);
                toks.push(range.token);
                i = range.end;
                continue;
            }

            // 普通单元格引用（可能是区域 A1:B2）
            if (CELL_RE.test(word)) {
                const range = readRangeTail(text, j, undefined, word.replace(/\$/g, ''));
                toks.push(range.token);
                i = range.end;
                continue;
            }

            // 定义名称：无法求值，按空处理
            toks.push({ kind: 'name', name: word });
            i = j;
            continue;
        }

        // 比较运算符
        if ((ch === '<' || ch === '>' || ch === '=') && text[i + 1] === '=') {
            toks.push({ kind: 'op', op: ch + '=' });
            i += 2;
            continue;
        }
        if (ch === '<' && text[i + 1] === '>') {
            toks.push({ kind: 'op', op: '<>' });
            i += 2;
            continue;
        }

        if ('+-*/^%&(),:;<>='.includes(ch)) {
            toks.push({ kind: 'op', op: ch });
            i++;
            continue;
        }

        // 其它字符：跳过
        i++;
    }

    return toks;
}

/** 读取引用后的 `:B9` 尾巴，产出 range / ref token */
function readRangeTail(
    text: string,
    start: number,
    sheet: string | undefined,
    firstRef: string
): { token: Token; end: number } {
    let i = start;
    while (i < text.length && isSpace(text[i])) i++;
    if (text[i] === ':') {
        const ref = readRefToken(text, i + 1);
        if (ref) {
            return {
                token: { kind: 'range', sheet, from: firstRef, to: ref.ref },
                end: ref.end
            };
        }
    }
    return { token: { kind: 'ref', sheet, ref: firstRef }, end: start };
}

// ---------- 求值 ----------

/** 区域值：保留二维结构，供按位置取数的函数使用 */
interface RangeValue {
    kind: 'range';
    rows: (FormulaValue | null)[][];
}

type Operand = FormulaValue | null | RangeValue;

function isRange(v: Operand): v is RangeValue {
    return typeof v === 'object' && v !== null && (v as RangeValue).kind === 'range';
}

/** 区域摊平成一维值列表 */
function flatten(v: Operand): (FormulaValue | null)[] {
    if (isRange(v)) {
        const out: (FormulaValue | null)[] = [];
        for (const row of v.rows) for (const cell of row) out.push(cell);
        return out;
    }
    return [v];
}

/** 数值化：文本按 Excel 规则解析（前导货币符号与千分位忽略） */
function toNum(v: FormulaValue | null): number {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const text = String(v).trim().replace(/[¥$€£,\s]/g, '');
    if (text === '') return 0;
    const n = parseFloat(text);
    return Number.isNaN(n) ? NaN : n;
}

/** 文本化 */
function toStr(v: FormulaValue | null): string {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return formatPlainNumber(v);
}

/** 数字转普通文本（不带科学计数法） */
function formatPlainNumber(n: number): string {
    if (!Number.isFinite(n)) return String(n);
    if (Number.isInteger(n)) return String(n);
    const abs = Math.abs(n);
    if (abs !== 0 && (abs < 1e-10 || abs >= 1e15)) return n.toExponential();
    return String(Number(n.toFixed(10)));
}

/** 布尔化（Excel：非零为真） */
function toBool(v: FormulaValue | null): boolean {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v === null) return false;
    const text = String(v).trim().toUpperCase();
    if (text === 'TRUE') return true;
    if (text === 'FALSE') return false;
    const n = parseFloat(text);
    return !Number.isNaN(n) && n !== 0;
}

/** 条件匹配器：支持 >、>=、<、<=、<>、=、通配符 * ? */
function criteriaMatcher(criteria: FormulaValue | null): (v: FormulaValue | null) => boolean {
    if (typeof criteria === 'number') {
        return (v) => toNum(v) === criteria;
    }
    if (typeof criteria === 'boolean') {
        return (v) => toBool(v) === criteria;
    }
    const text = String(criteria ?? '').trim();
    const m = /^(>=|<=|<>|>|<|=)?(.*)$/.exec(text);
    const op = m && m[1] ? m[1] : '=';
    const rest = m ? m[2].trim() : text;
    const rhs = rest === '' ? NaN : Number(rest);
    const numeric = !Number.isNaN(rhs) && rest !== '';

    if (numeric) {
        return (v) => {
            const n = toNum(v);
            if (Number.isNaN(n)) return false;
            switch (op) {
                case '>': return n > rhs;
                case '<': return n < rhs;
                case '>=': return n >= rhs;
                case '<=': return n <= rhs;
                case '<>': return n !== rhs;
                default: return n === rhs;
            }
        };
    }

    // 文本：支持 * ? 通配
    const hasWildcard = /[*?]/.test(rest);
    const pattern = hasWildcard
        ? new RegExp(
              '^' +
                  rest
                      .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                      .replace(/\\\*/g, '.*')
                      .replace(/\\\?/g, '.') +
                  '$',
              'i'
          )
        : null;

    return (v) => {
        if (op === '<>') {
            if (v === null || v === undefined || v === '') return rest !== '';
            return toStr(v).toLowerCase() !== rest.toLowerCase();
        }
        const text2 = v === null || v === undefined ? '' : toStr(v);
        if (pattern) return pattern.test(text2);
        return text2.toLowerCase() === rest.toLowerCase();
    };
}

/** Excel 日期序列号 -> Date */
function serialToDate(serial: number): Date {
    return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

export { formatByPattern };

/** 求值入口 */
export function evaluateFormula(formula: string, ctx: FormulaContext): FormulaValue {
    const toks = tokenizeFormula(formula);
    if (!toks.length) return 0;
    const parser = new Parser(toks, ctx);
    const result = parser.parseExpression();
    if (isRange(result)) {
        const flat = flatten(result);
        return flat.length ? (flat[0] ?? 0) : 0;
    }
    return result ?? 0;
}

/** 解析器 */
class Parser {
    private pos = 0;

    constructor(private readonly toks: Token[], private readonly ctx: FormulaContext) {}

    private peek(): Token | undefined {
        return this.toks[this.pos];
    }

    private eat(op: string): boolean {
        const t = this.peek();
        if (t && t.kind === 'op' && t.op === op) {
            this.pos++;
            return true;
        }
        return false;
    }

    private expect(op: string): void {
        if (!this.eat(op)) throw new FormulaError('#VALUE!');
    }

    parseExpression(): Operand {
        return this.parseCompare();
    }

    private parseCompare(): Operand {
        let left = this.parseConcat();
        for (;;) {
            const t = this.peek();
            if (!t || t.kind !== 'op' || !['=', '<>', '<', '>', '<=', '>='].includes(t.op)) break;
            this.pos++;
            const right = this.parseConcat();
            left = compareValues(left, right, t.op);
        }
        return left;
    }

    private parseConcat(): Operand {
        let left = this.parseAddSub();
        while (this.peek()?.kind === 'op' && (this.peek() as { op: string }).op === '&') {
            this.pos++;
            const right = this.parseAddSub();
            left = toStr(scalar(left)) + toStr(scalar(right));
        }
        return left;
    }

    private parseAddSub(): Operand {
        let left = this.parseMulDiv();
        for (;;) {
            const t = this.peek();
            if (!t || t.kind !== 'op' || (t.op !== '+' && t.op !== '-')) break;
            this.pos++;
            const right = this.parseMulDiv();
            const a = toNum(scalar(left));
            const b = toNum(scalar(right));
            // 「文本 + 数字」在 Excel 里是 #VALUE!，但空值按 0 处理
            if (Number.isNaN(a) || Number.isNaN(b)) throw new FormulaError('#VALUE!');
            left = t.op === '+' ? a + b : a - b;
        }
        return left;
    }

    private parseMulDiv(): Operand {
        let left = this.parsePower();
        for (;;) {
            const t = this.peek();
            if (!t || t.kind !== 'op' || (t.op !== '*' && t.op !== '/')) break;
            this.pos++;
            const right = this.parsePower();
            const a = toNum(scalar(left));
            const b = toNum(scalar(right));
            if (Number.isNaN(a) || Number.isNaN(b)) throw new FormulaError('#VALUE!');
            if (t.op === '*') left = a * b;
            else {
                if (b === 0) throw new FormulaError('#DIV/0!');
                left = a / b;
            }
        }
        return left;
    }

    private parsePower(): Operand {
        let left = this.parseUnary();
        while (this.peek()?.kind === 'op' && (this.peek() as { op: string }).op === '^') {
            this.pos++;
            const right = this.parseUnary();
            left = Math.pow(toNum(scalar(left)), toNum(scalar(right)));
        }
        return left;
    }

    private parseUnary(): Operand {
        const t = this.peek();
        if (t && t.kind === 'op' && (t.op === '-' || t.op === '+')) {
            this.pos++;
            const v = this.parseUnary();
            const n = toNum(scalar(v));
            return t.op === '-' ? -n : n;
        }
        return this.parsePostfix();
    }

    private parsePostfix(): Operand {
        let v = this.parsePrimary();
        while (this.peek()?.kind === 'op' && (this.peek() as { op: string }).op === '%') {
            this.pos++;
            v = toNum(scalar(v)) / 100;
        }
        return v;
    }

    private parsePrimary(): Operand {
        const t = this.peek();
        if (!t) throw new FormulaError('#VALUE!');

        if (t.kind === 'num') {
            this.pos++;
            return t.value;
        }
        if (t.kind === 'str') {
            this.pos++;
            return t.value;
        }
        if (t.kind === 'bool') {
            this.pos++;
            return t.value;
        }
        if (t.kind === 'name') {
            this.pos++;
            return 0;
        }
        if (t.kind === 'func') {
            this.pos++;
            this.expect('(');
            const args: Operand[] = [];
            if (!(this.peek()?.kind === 'op' && (this.peek() as { op: string }).op === ')')) {
                args.push(this.parseExpression());
                while (this.eat(',')) args.push(this.parseExpression());
            }
            this.expect(')');
            return callFunction(t.name, args, this.ctx);
        }
        if (t.kind === 'ref') {
            this.pos++;
            return this.readRefValue(t.sheet, t.ref);
        }
        if (t.kind === 'range') {
            this.pos++;
            return this.readRangeValue(t.sheet, t.from, t.to);
        }
        if (t.kind === 'op' && t.op === '(') {
            this.pos++;
            const v = this.parseExpression();
            this.expect(')');
            return v;
        }
        throw new FormulaError('#VALUE!');
    }

    /** 读取单元格值 */
    private readRefValue(sheet: string | undefined, ref: string): Operand {
        const target = this.resolveSheet(sheet);
        const pos = parseCellRef(ref);
        if (!pos) throw new FormulaError('#REF!');
        const cell = this.ctx.read(target, pos.row, pos.col);
        if (cell.value !== undefined && cell.value !== null) return cell.value;
        if (cell.text !== undefined && cell.text !== '') return cell.text;
        return 0;
    }

    /** 读取区域值 */
    private readRangeValue(sheet: string | undefined, from: string, to: string): Operand {
        const target = this.resolveSheet(sheet);
        const a = parseCellRef(from);
        const b = parseCellRef(to);
        if (!a || !b) throw new FormulaError('#REF!');
        const r1 = Math.min(a.row, b.row);
        const r2 = Math.max(a.row, b.row);
        const c1 = Math.min(a.col, b.col);
        const c2 = Math.max(a.col, b.col);
        const rows: (FormulaValue | null)[][] = [];
        for (let r = r1; r <= r2; r++) {
            const row: (FormulaValue | null)[] = [];
            for (let c = c1; c <= c2; c++) {
                const cell = this.ctx.read(target, r, c);
                row.push(cell.value ?? null);
            }
            rows.push(row);
        }
        return { kind: 'range', rows };
    }

    /** 解析表名（相对引用落在当前表） */
    private resolveSheet(sheet: string | undefined): string {
        const name = sheet || this.ctx.sheetName;
        if (this.ctx.hasSheet && !this.ctx.hasSheet(name)) throw new FormulaError('#REF!');
        return name;
    }
}

/** 取标量值（区域取首个单元格） */
function scalar(v: Operand): FormulaValue | null {
    if (isRange(v)) {
        const flat = flatten(v);
        return flat.length ? flat[0] ?? null : null;
    }
    return v;
}

/** 比较 */
function compareValues(a: Operand, b: Operand, op: string): boolean {
    const av = scalar(a);
    const bv = scalar(b);
    let cmp: number;
    const an = toNum(av);
    const bn = toNum(bv);
    if (!Number.isNaN(an) && !Number.isNaN(bn) && typeof av !== 'string' && typeof bv !== 'string') {
        cmp = an === bn ? 0 : an < bn ? -1 : 1;
    } else {
        const as = toStr(av).toLowerCase();
        const bs = toStr(bv).toLowerCase();
        cmp = as === bs ? 0 : as < bs ? -1 : 1;
    }
    switch (op) {
        case '=': return cmp === 0;
        case '<>': return cmp !== 0;
        case '<': return cmp < 0;
        case '>': return cmp > 0;
        case '<=': return cmp <= 0;
        default: return cmp >= 0;
    }
}

/** 数值集合（忽略文本/布尔） */
function numbersOf(v: Operand): number[] {
    return flatten(v)
        .map((x) => toNum(x))
        .filter((n) => !Number.isNaN(n));
}

/** 函数实现 */
function callFunction(name: string, args: Operand[], ctx: FormulaContext): Operand {
    const num = (i: number) => toNum(scalar(args[i]));
    const str = (i: number) => toStr(scalar(args[i]));
    const nums = (): number[] => args.flatMap((a) => numbersOf(a));

    switch (name) {
        // ---- 数学 ----
        case 'SUM': return nums().reduce((s, n) => s + n, 0);
        case 'PRODUCT': {
            const list = nums();
            return list.length ? list.reduce((s, n) => s * n, 1) : 0;
        }
        case 'ABS': return Math.abs(num(0));
        case 'SQRT': return Math.sqrt(num(0));
        case 'POWER': return Math.pow(num(0), num(1));
        case 'MOD': {
            const b = num(1);
            if (b === 0) throw new FormulaError('#DIV/0!');
            return num(0) % b;
        }
        case 'INT': return Math.floor(num(0));
        case 'TRUNC': return Math.trunc(num(0));
        case 'ROUND': {
            const d = args.length > 1 ? num(1) : 0;
            const f = Math.pow(10, d);
            const v = num(0) * f;
            // 四舍五入（远离零方向修正，避免 .5 被银行家舍入影响）
            const r = Math.sign(v) * Math.round(Math.abs(v) + Number.EPSILON * Math.abs(v));
            return r / f;
        }
        case 'ROUNDUP': {
            const d = args.length > 1 ? num(1) : 0;
            const f = Math.pow(10, d);
            return Math.ceil(num(0) * f) / f;
        }
        case 'ROUNDDOWN': {
            const d = args.length > 1 ? num(1) : 0;
            const f = Math.pow(10, d);
            return Math.floor(num(0) * f) / f;
        }
        case 'CEILING': return Math.ceil(num(0));
        case 'FLOOR': return Math.floor(num(0));
        case 'SIGN': return Math.sign(num(0));
        case 'EXP': return Math.exp(num(0));
        case 'LN': return Math.log(num(0));
        case 'LOG': return Math.log(num(0)) / Math.log(args.length > 1 ? num(1) : 10);
        case 'LOG10': return Math.log10(num(0));
        case 'PI': return Math.PI;
        case 'SUMPRODUCT': {
            const cols = args.map((a) => numbersOf(a));
            if (!cols.length) return 0;
            const len = Math.min(...cols.map((c) => c.length));
            let total = 0;
            for (let i = 0; i < len; i++) {
                let p = 1;
                for (const c of cols) p *= c[i] ?? 0;
                total += p;
            }
            return total;
        }

        // ---- 统计 ----
        case 'AVERAGE': {
            const list = nums();
            return list.length ? list.reduce((s, n) => s + n, 0) / list.length : 0;
        }
        case 'COUNT': return nums().length;
        case 'COUNTA': {
            let n = 0;
            for (const a of args) {
                for (const v of flatten(a)) if (v !== null && v !== undefined && v !== '') n++;
            }
            return n;
        }
        case 'COUNTBLANK': {
            let n = 0;
            for (const a of args) {
                for (const v of flatten(a)) if (v === null || v === undefined || v === '') n++;
            }
            return n;
        }
        case 'MAX': {
            const list = nums();
            return list.length ? Math.max(...list) : 0;
        }
        case 'MIN': {
            const list = nums();
            return list.length ? Math.min(...list) : 0;
        }
        case 'MEDIAN': {
            const list = nums().sort((a, b) => a - b);
            if (!list.length) return 0;
            const mid = Math.floor(list.length / 2);
            return list.length % 2 ? list[mid] : (list[mid - 1] + list[mid]) / 2;
        }
        case 'LARGE': {
            const list = nums().sort((a, b) => b - a);
            const k = Math.max(1, Math.floor(num(1)));
            return list[k - 1] ?? 0;
        }
        case 'SMALL': {
            const list = nums().sort((a, b) => a - b);
            const k = Math.max(1, Math.floor(num(1)));
            return list[k - 1] ?? 0;
        }
        case 'RANK': {
            const list = nums();
            const target = num(0);
            const desc = args.length > 2 ? toBool(scalar(args[2])) : false;
            const sorted = [...list].sort((a, b) => (desc ? b - a : a - b));
            return sorted.indexOf(target) + 1;
        }
        case 'COUNTIF': {
            const match = criteriaMatcher(scalar(args[1]));
            return flatten(args[0]).filter((v) => match(v)).length;
        }
        case 'COUNTIFS': {
            const pairs: { match: (v: FormulaValue | null) => boolean; range: Operand }[] = [];
            for (let i = 0; i + 1 < args.length; i += 2) {
                pairs.push({ range: args[i], match: criteriaMatcher(scalar(args[i + 1])) });
            }
            if (!pairs.length) return 0;
            const base = flatten(pairs[0].range);
            let n = 0;
            for (let i = 0; i < base.length; i++) {
                if (pairs.every((p) => p.match(flatten(p.range)[i] ?? null))) n++;
            }
            return n;
        }
        case 'SUMIF': {
            const match = criteriaMatcher(scalar(args[1]));
            const src = flatten(args[0]);
            const sumSrc = args.length > 2 ? flatten(args[2]) : src;
            let total = 0;
            for (let i = 0; i < src.length; i++) {
                if (!match(src[i] ?? null)) continue;
                const v = toNum(sumSrc[i] ?? null);
                if (!Number.isNaN(v)) total += v;
            }
            return total;
        }
        case 'SUMIFS': {
            const sumSrc = flatten(args[0]);
            const pairs: { match: (v: FormulaValue | null) => boolean; range: Operand }[] = [];
            for (let i = 1; i + 1 < args.length; i += 2) {
                pairs.push({ range: args[i], match: criteriaMatcher(scalar(args[i + 1])) });
            }
            let total = 0;
            for (let i = 0; i < sumSrc.length; i++) {
                if (pairs.every((p) => p.match(flatten(p.range)[i] ?? null))) {
                    const v = toNum(sumSrc[i]);
                    if (!Number.isNaN(v)) total += v;
                }
            }
            return total;
        }
        case 'AVERAGEIF': {
            const match = criteriaMatcher(scalar(args[1]));
            const src = flatten(args[0]);
            const avgSrc = args.length > 2 ? flatten(args[2]) : src;
            const picked: number[] = [];
            for (let i = 0; i < src.length; i++) {
                if (!match(src[i] ?? null)) continue;
                const v = toNum(avgSrc[i] ?? null);
                if (!Number.isNaN(v)) picked.push(v);
            }
            return picked.length ? picked.reduce((s, n) => s + n, 0) / picked.length : 0;
        }

        // ---- 逻辑 ----
        case 'IF': {
            const cond = toBool(scalar(args[0]));
            if (cond) return args.length > 1 ? scalar(args[1]) ?? 0 : 0;
            return args.length > 2 ? scalar(args[2]) ?? 0 : 0;
        }
        case 'IFS': {
            for (let i = 0; i + 1 < args.length; i += 2) {
                if (toBool(scalar(args[i]))) return scalar(args[i + 1]) ?? 0;
            }
            throw new FormulaError('#N/A');
        }
        case 'AND': return args.every((a) => toBool(scalar(a))) ? 1 : 0;
        case 'OR': return args.some((a) => toBool(scalar(a))) ? 1 : 0;
        case 'NOT': return toBool(scalar(args[0])) ? 0 : 1;
        case 'XOR': {
            let n = 0;
            for (const a of args) if (toBool(scalar(a))) n++;
            return n % 2 === 1 ? 1 : 0;
        }
        case 'TRUE': return true;
        case 'FALSE': return false;
        case 'IFERROR': {
            try {
                const v = scalar(args[0]);
                if (v === null) return scalar(args[1]) ?? 0;
                return v;
            } catch {
                return scalar(args[1]) ?? 0;
            }
        }
        case 'IFNA': {
            try {
                return scalar(args[0]) ?? 0;
            } catch {
                return scalar(args[1]) ?? 0;
            }
        }
        case 'ISBLANK': {
            const v = scalar(args[0]);
            return v === null || v === undefined || v === '' ? 1 : 0;
        }
        case 'ISNUMBER': return typeof scalar(args[0]) === 'number' ? 1 : 0;
        case 'ISTEXT': return typeof scalar(args[0]) === 'string' ? 1 : 0;
        case 'ISLOGICAL': return typeof scalar(args[0]) === 'boolean' ? 1 : 0;
        case 'N': {
            const v = scalar(args[0]);
            return typeof v === 'number' ? v : 0;
        }

        // ---- 文本 ----
        case 'CONCAT':
        case 'CONCATENATE': {
            const parts: string[] = [];
            for (const a of args) for (const v of flatten(a)) parts.push(toStr(v ?? null));
            return parts.join('');
        }
        case 'TEXTJOIN': {
            const delim = str(0);
            const skipEmpty = toBool(scalar(args[1]));
            const parts: string[] = [];
            for (let i = 2; i < args.length; i++) {
                for (const v of flatten(args[i])) {
                    const text = toStr(v ?? null);
                    if (skipEmpty && text === '') continue;
                    parts.push(text);
                }
            }
            return parts.join(delim);
        }
        case 'LEFT': return str(0).slice(0, Math.max(0, Math.floor(num(1))));
        case 'RIGHT': {
            const s = str(0);
            const n = Math.max(0, Math.floor(num(1)));
            return n ? s.slice(-n) : '';
        }
        case 'MID': {
            const s = str(0);
            const st = Math.max(1, Math.floor(num(1)));
            const ln = Math.max(0, Math.floor(num(2)));
            return s.slice(st - 1, st - 1 + ln);
        }
        case 'LEN': return str(0).length;
        case 'LOWER': return str(0).toLowerCase();
        case 'UPPER': return str(0).toUpperCase();
        case 'TRIM': return str(0).trim().replace(/\s+/g, ' ');
        case 'PROPER': return str(0).replace(/\w\S*/g, (t) => t[0].toUpperCase() + t.slice(1).toLowerCase());
        case 'REPT': return str(0).repeat(Math.max(0, Math.floor(num(1))));
        case 'SUBSTITUTE': {
            const s = str(0);
            const from = str(1);
            const to = str(2);
            if (from === '') return s;
            if (args.length > 3) {
                const which = Math.max(1, Math.floor(num(3)));
                let index = -1;
                for (let k = 0; k < which; k++) index = s.indexOf(from, index + 1);
                if (index < 0) return s;
                return s.slice(0, index) + to + s.slice(index + from.length);
            }
            return s.split(from).join(to);
        }
        case 'REPLACE': {
            const s = str(0);
            const st = Math.max(1, Math.floor(num(1)));
            const len = Math.max(0, Math.floor(num(2)));
            return s.slice(0, st - 1) + str(3) + s.slice(st - 1 + len);
        }
        case 'FIND': return str(0).indexOf(str(1)) + 1;
        case 'SEARCH': {
            const hay = str(1).toLowerCase();
            const needle = str(0).toLowerCase();
            return hay.indexOf(needle) + 1;
        }
        case 'EXACT': return str(0) === str(1) ? 1 : 0;
        case 'TEXT': {
            const v = scalar(args[0]);
            if (v === null) return '';
            if (typeof v === 'number') return formatByPattern(v, str(1));
            return toStr(v);
        }
        case 'VALUE': {
            const n = parseFloat(String(scalar(args[0]) ?? '').replace(/[¥$€£,\s]/g, ''));
            if (Number.isNaN(n)) throw new FormulaError('#VALUE!');
            return n;
        }

        // ---- 查找 ----
        case 'VLOOKUP':
        case 'HLOOKUP': {
            const lookup = scalar(args[0]);
            const table = args[1];
            if (!isRange(table)) throw new FormulaError('#VALUE!');
            const colIndex = Math.floor(num(2));
            const approx = args.length > 3 ? toBool(scalar(args[3])) : true;
            const rows = name === 'VLOOKUP' ? table.rows : null;
            if (!rows) throw new FormulaError('#VALUE!');
            let found: Operand = null;
            let best: number | null = null;
            for (const row of rows) {
                const key = row[0] ?? null;
                const eq = compareValues(key ?? null, lookup ?? null, '=');
                if (eq) {
                    found = row[colIndex - 1] ?? null;
                    best = 0;
                    break;
                }
                if (approx && typeof key === 'number' && typeof lookup === 'number' && key <= lookup) {
                    best = key;
                    found = row[colIndex - 1] ?? null;
                }
            }
            if (found === null && best === null) throw new FormulaError('#N/A');
            return found ?? 0;
        }
        case 'INDEX': {
            const table = args[0];
            if (!isRange(table)) throw new FormulaError('#VALUE!');
            const r = Math.floor(num(1));
            const c = args.length > 2 ? Math.floor(num(2)) : 1;
            if (r < 1 || c < 1) throw new FormulaError('#VALUE!');
            return table.rows[r - 1]?.[c - 1] ?? 0;
        }
        case 'MATCH': {
            const lookup = scalar(args[0]);
            const list = flatten(args[1]);
            const type = args.length > 2 ? Math.floor(num(2)) : 1;
            if (type === 0) {
                const match = criteriaMatcher(lookup);
                for (let i = 0; i < list.length; i++) if (match(list[i] ?? null)) return i + 1;
                throw new FormulaError('#N/A');
            }
            const target = toNum(lookup);
            let best = -1;
            for (let i = 0; i < list.length; i++) {
                const v = toNum(list[i] ?? null);
                if (Number.isNaN(v)) continue;
                if (type === 1 && v <= target) best = i;
                if (type === -1 && v >= target) best = i;
            }
            if (best < 0) throw new FormulaError('#N/A');
            return best + 1;
        }
        case 'CHOOSE': {
            const idx = Math.floor(num(0));
            if (idx < 1 || idx >= args.length) throw new FormulaError('#VALUE!');
            return scalar(args[idx]) ?? 0;
        }
        case 'ROW': return 0;
        case 'COLUMN': return 0;

        // ---- 日期 ----
        case 'TODAY': {
            const now = new Date();
            const utc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
            return Math.floor(utc / 86400000 + 25569);
        }
        case 'NOW': return Date.now() / 86400000 + 25569;
        case 'DATE': {
            const y = Math.floor(num(0));
            const m = Math.floor(num(1));
            const d = Math.floor(num(2));
            const ms = Date.UTC(y, m - 1, d);
            return Math.floor(ms / 86400000 + 25569);
        }
        case 'YEAR': return serialToDate(num(0)).getUTCFullYear();
        case 'MONTH': return serialToDate(num(0)).getUTCMonth() + 1;
        case 'DAY': return serialToDate(num(0)).getUTCDate();
        case 'HOUR': return serialToDate(num(0)).getUTCHours();
        case 'MINUTE': return serialToDate(num(0)).getUTCMinutes();
        case 'SECOND': return serialToDate(num(0)).getUTCSeconds();

        default:
            throw new FormulaError('#NAME?');
    }
}

/**
 * 提取公式中的引用（用于依赖追踪）
 *
 * 返回的 ref 形如 `B3` / `B3:C9`，sheet 为跨表时的表名（未带引号）。
 */
export function extractFormulaRefs(formula: string): { sheet?: string; ref: string }[] {
    let toks: Token[];
    try {
        toks = tokenizeFormula(formula);
    } catch {
        return [];
    }
    const out: { sheet?: string; ref: string }[] = [];
    for (const t of toks) {
        if (t.kind === 'ref' && t.ref) out.push({ sheet: t.sheet, ref: t.ref });
        else if (t.kind === 'range' && t.from && t.to) out.push({ sheet: t.sheet, ref: `${t.from}:${t.to}` });
    }
    return out;
}
