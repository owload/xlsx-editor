import { decodeCol, encodeCol } from './addr';
import { formatNumber } from './numfmt';
import type { Book } from './types';
import { usedSize } from './types';

export interface FErr {
  err: string;
}
export type Val = number | string | boolean | null | FErr;
interface RangeVal {
  range: Val[];
}
type Arg = Val | RangeVal;

const E = (err: string): FErr => ({ err });
export const isErr = (v: unknown): v is FErr => typeof v === 'object' && v !== null && 'err' in v;
const isRange = (a: Arg): a is RangeVal => typeof a === 'object' && a !== null && 'range' in a;

export const NUM_RE = /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i;
const ERR_RE = /^#(DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|CIRC!|ERROR!)$/;

export function parseLiteral(raw: string, text?: boolean): Val {
  if (raw === '') return null;
  if (text) return raw;
  const t = raw.trim();
  if (NUM_RE.test(t)) return Number(t);
  const u = t.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  if (ERR_RE.test(t)) return E(t);
  return raw;
}

// ---------- cell references ----------

interface Pt {
  c: number;
  r: number;
  ac: boolean;
  ar: boolean;
}
export interface RefParts {
  sheet: string | null;
  a: Pt;
  b?: Pt;
}

const REF_SRC = String.raw`(?<![\p{L}\p{N}_.$'])((?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![\p{L}\p{N}_(!])`;

function partsFrom(g: ArrayLike<string | undefined>): RefParts {
  let sheet: string | null = null;
  if (g[1]) {
    const s = g[1].slice(0, -1);
    sheet = s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s;
  }
  const a: Pt = { ac: !!g[2], c: decodeCol(g[3]!.toUpperCase()), ar: !!g[4], r: Number(g[5]) - 1 };
  const b: Pt | undefined = g[7]
    ? { ac: !!g[6], c: decodeCol(g[7].toUpperCase()), ar: !!g[8], r: Number(g[9]) - 1 }
    : undefined;
  return { sheet, a, b };
}

const quoteSheet = (n: string) =>
  /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;

const fmtPt = (p: Pt) => `${p.ac ? '$' : ''}${encodeCol(p.c)}${p.ar ? '$' : ''}${p.r + 1}`;

export function formatRef(p: RefParts): string {
  return (p.sheet ? quoteSheet(p.sheet) + '!' : '') + fmtPt(p.a) + (p.b ? ':' + fmtPt(p.b) : '');
}

/** Rewrites every reference in a formula, leaving string literals alone. */
export function mapRefs(formula: string, fn: (p: RefParts) => string): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((seg, i) => (i % 2 ? seg : seg.replace(new RegExp(REF_SRC, 'gu'), (...g) => fn(partsFrom(g)))))
    .join('');
}

/**
 * Adjusts a formula when rows/columns of sheet `target` are inserted
 * (delta > 0) or deleted (delta < 0) starting at index `at`.
 */
export function adjustForEdit(
  raw: string,
  own: string,
  target: string,
  axis: 'r' | 'c',
  at: number,
  delta: number,
): string {
  if (raw[0] !== '=') return raw;
  const tl = target.toLowerCase();
  return (
    '=' +
    mapRefs(raw.slice(1), (p) => {
      if ((p.sheet ?? own).toLowerCase() !== tl) return formatRef(p);
      const a = { ...p.a };
      const b = p.b && { ...p.b };
      if (delta > 0) {
        if (a[axis] >= at) a[axis] += delta;
        if (b && b[axis] >= at) b[axis] += delta;
      } else {
        const n = -delta;
        const end = at + n;
        if (!b) {
          if (a[axis] >= at && a[axis] < end) return '#REF!';
          if (a[axis] >= end) a[axis] -= n;
        } else {
          const s = a[axis];
          const e = b[axis];
          const ns = s < at ? s : s >= end ? s - n : at;
          const ne = e < at ? e : e >= end ? e - n : at - 1;
          if (ne < ns) return '#REF!';
          a[axis] = ns;
          b[axis] = ne;
        }
      }
      return formatRef({ ...p, a, b });
    })
  );
}

/** Shifts relative references (as when filling or copying a formula). */
export function shiftFormula(raw: string, dr: number, dc: number): string {
  if (raw[0] !== '=') return raw;
  return (
    '=' +
    mapRefs(raw.slice(1), (p) => {
      const mv = (pt: Pt): Pt | null => {
        const r = pt.ar ? pt.r : pt.r + dr;
        const c = pt.ac ? pt.c : pt.c + dc;
        return r < 0 || c < 0 || r > 1048575 || c > 16383 ? null : { ...pt, r, c };
      };
      const a = mv(p.a);
      const b = p.b ? mv(p.b) : undefined;
      if (!a || b === null) return '#REF!';
      return formatRef({ ...p, a, b });
    })
  );
}

export function renameSheetInFormula(raw: string, from: string, to: string): string {
  if (raw[0] !== '=') return raw;
  const fl = from.toLowerCase();
  return '=' + mapRefs(raw.slice(1), (p) => formatRef(p.sheet && p.sheet.toLowerCase() === fl ? { ...p, sheet: to } : p));
}

// ---------- tokenizer and parser ----------

type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; p: RefParts }
  | { t: 'id'; v: string }
  | { t: 'op'; v: string };

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  const refRe = new RegExp(REF_SRC, 'yu');
  const numRe = /(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/y;
  const idRe = /[\p{L}_][\p{L}\p{N}_.]*/yu;
  const opRe = /<>|<=|>=|[-+*/^&=<>%(),;]/y;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      for (;;) {
        if (j >= src.length) throw new Error('unterminated string');
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j++];
      }
      out.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    numRe.lastIndex = i;
    let m = numRe.exec(src);
    if (m) {
      out.push({ t: 'num', v: Number(m[0]) });
      i += m[0].length;
      continue;
    }
    refRe.lastIndex = i;
    m = refRe.exec(src);
    if (m) {
      out.push({ t: 'ref', p: partsFrom(m) });
      i += m[0].length;
      continue;
    }
    idRe.lastIndex = i;
    m = idRe.exec(src);
    if (m) {
      out.push({ t: 'id', v: m[0] });
      i += m[0].length;
      continue;
    }
    opRe.lastIndex = i;
    m = opRe.exec(src);
    if (m) {
      out.push({ t: 'op', v: m[0] === ';' ? ',' : m[0] });
      i += m[0].length;
      continue;
    }
    throw new Error('bad char ' + ch);
  }
  return out;
}

type Node =
  | { k: 'lit'; v: Val }
  | { k: 'ref'; p: RefParts }
  | { k: 'un'; op: string; a: Node }
  | { k: 'pct'; a: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'call'; name: string; args: Node[] };

function parse(src: string): Node {
  const toks = tokenize(src);
  let pos = 0;
  const peekOp = (...ops: string[]) => {
    const t = toks[pos];
    return t && t.t === 'op' && ops.includes(t.v) ? t.v : null;
  };
  const level = (next: () => Node, ...ops: string[]): Node => {
    let a = next();
    for (let op = peekOp(...ops); op; op = peekOp(...ops)) {
      pos++;
      a = { k: 'bin', op, a, b: next() };
    }
    return a;
  };
  const cmp = (): Node => level(concat, '=', '<>', '<', '>', '<=', '>=');
  const concat = (): Node => level(add, '&');
  const add = (): Node => level(mul, '+', '-');
  const mul = (): Node => level(pow, '*', '/');
  const pow = (): Node => level(unary, '^');
  const unary = (): Node => {
    const op = peekOp('-', '+');
    if (op) {
      pos++;
      return { k: 'un', op, a: unary() };
    }
    return postfix();
  };
  const postfix = (): Node => {
    let a = primary();
    while (peekOp('%')) {
      pos++;
      a = { k: 'pct', a };
    }
    return a;
  };
  const primary = (): Node => {
    const t = toks[pos++];
    if (!t) throw new Error('unexpected end');
    if (t.t === 'num') return { k: 'lit', v: t.v };
    if (t.t === 'str') return { k: 'lit', v: t.v };
    if (t.t === 'ref') return { k: 'ref', p: t.p };
    if (t.t === 'id') {
      if (peekOp('(')) {
        pos++;
        const args: Node[] = [];
        if (!peekOp(')')) {
          do {
            args.push(cmp());
          } while (peekOp(',') && ++pos);
        }
        if (!peekOp(')')) throw new Error('expected )');
        pos++;
        return { k: 'call', name: t.v.toUpperCase(), args };
      }
      const u = t.v.toUpperCase();
      if (u === 'TRUE' || u === 'FALSE') return { k: 'lit', v: u === 'TRUE' };
      return { k: 'lit', v: E('#NAME?') };
    }
    if (t.v === '(') {
      const n = cmp();
      if (!peekOp(')')) throw new Error('expected )');
      pos++;
      return n;
    }
    throw new Error('unexpected ' + t.v);
  };
  const node = cmp();
  if (pos < toks.length) throw new Error('trailing tokens');
  return node;
}

const parseCache = new Map<string, Node | null>();
function parseCached(src: string): Node | null {
  let n = parseCache.get(src);
  if (n === undefined) {
    try {
      n = parse(src);
    } catch {
      n = null;
    }
    if (parseCache.size > 5000) parseCache.clear();
    parseCache.set(src, n);
  }
  return n;
}

// ---------- evaluation ----------

const toNum = (v: Val): number | FErr => {
  if (isErr(v)) return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'number') return v;
  const t = v.trim();
  if (t === '') return 0;
  return NUM_RE.test(t) ? Number(t) : E('#VALUE!');
};
const toStr = (v: Val): string => (v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
const toBool = (v: Val): boolean | FErr => {
  if (isErr(v)) return v;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return E('#VALUE!');
  }
  return !!toNum(v);
};

const rank = (v: Val) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
function compare(a: Val, b: Val): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') {
    const x = a.toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  const x = Number(a);
  const y = Number(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

export function createEvaluator(book: Book) {
  const cache = new Map<string, Val>();
  const busy = new Set<string>();
  const sizes = book.sheets.map(usedSize);
  const sheetIdx = (name: string) => book.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());

  function value(si: number, r: number, c: number): Val {
    const cell = book.sheets[si]?.rows[r]?.[c];
    if (!cell || cell.raw === '') return null;
    if (cell.raw[0] !== '=' || cell.text) return parseLiteral(cell.raw, cell.text);
    const key = `${si}:${r}:${c}`;
    const hit = cache.get(key);
    if (hit !== undefined) return hit;
    if (busy.has(key)) return E('#CIRC!');
    busy.add(key);
    let res: Val;
    const ast = parseCached(cell.raw.slice(1));
    if (!ast) res = E('#ERROR!');
    else {
      const out = evalNode(ast, si);
      res = isRange(out) ? E('#VALUE!') : out === null ? 0 : out;
    }
    busy.delete(key);
    cache.set(key, res);
    return res;
  }

  const scalar = (n: Node, si: number): Val => {
    const v = evalNode(n, si);
    if (!isRange(v)) return v;
    return v.range.length === 1 ? v.range[0] : E('#VALUE!');
  };

  function evalNode(n: Node, si: number): Arg {
    switch (n.k) {
      case 'lit':
        return n.v;
      case 'ref': {
        const ti = n.p.sheet ? sheetIdx(n.p.sheet) : si;
        if (ti < 0) return E('#REF!');
        if (!n.p.b) return value(ti, n.p.a.r, n.p.a.c);
        const r1 = Math.min(n.p.a.r, n.p.b.r);
        const r2 = Math.min(Math.max(n.p.a.r, n.p.b.r), sizes[ti].rows - 1);
        const c1 = Math.min(n.p.a.c, n.p.b.c);
        const c2 = Math.min(Math.max(n.p.a.c, n.p.b.c), sizes[ti].cols - 1);
        const range: Val[] = [];
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) range.push(value(ti, r, c));
        return { range };
      }
      case 'un': {
        const v = toNum(scalar(n.a, si));
        return isErr(v) ? v : n.op === '-' ? -v : v;
      }
      case 'pct': {
        const v = toNum(scalar(n.a, si));
        return isErr(v) ? v : v / 100;
      }
      case 'bin': {
        const a = scalar(n.a, si);
        const b = scalar(n.b, si);
        if (isErr(a)) return a;
        if (isErr(b)) return b;
        switch (n.op) {
          case '&':
            return toStr(a) + toStr(b);
          case '=':
            return compare(a, b) === 0;
          case '<>':
            return compare(a, b) !== 0;
          case '<':
            return compare(a, b) < 0;
          case '>':
            return compare(a, b) > 0;
          case '<=':
            return compare(a, b) <= 0;
          case '>=':
            return compare(a, b) >= 0;
        }
        const x = toNum(a);
        const y = toNum(b);
        if (isErr(x)) return x;
        if (isErr(y)) return y;
        switch (n.op) {
          case '+':
            return x + y;
          case '-':
            return x - y;
          case '*':
            return x * y;
          case '/':
            return y === 0 ? E('#DIV/0!') : x / y;
          default: {
            const p = Math.pow(x, y);
            return Number.isFinite(p) ? p : E('#NUM!');
          }
        }
      }
      case 'call':
        return call(n, si);
    }
  }

  /** Numbers from arguments: only numbers count inside ranges. */
  function nums(args: Arg[]): number[] | FErr {
    const out: number[] = [];
    for (const a of args) {
      if (isRange(a)) {
        for (const v of a.range) {
          if (isErr(v)) return v;
          if (typeof v === 'number') out.push(v);
        }
      } else {
        if (a === null) continue;
        const x = toNum(a);
        if (isErr(x)) return x;
        out.push(x);
      }
    }
    return out;
  }

  function call(n: Extract<Node, { k: 'call' }>, si: number): Arg {
    const name = n.name;
    if (name === 'IF') {
      const c = toBool(scalar(n.args[0], si));
      if (isErr(c)) return c;
      const branch = c ? n.args[1] : n.args[2];
      return branch ? scalar(branch, si) : c;
    }
    if (name === 'IFERROR') {
      const v = scalar(n.args[0], si);
      return isErr(v) ? scalar(n.args[1], si) : v;
    }
    const args = n.args.map((a) => evalNode(a, si));
    const sc = (i: number): Val => {
      const a = args[i];
      if (a === undefined) return null;
      return isRange(a) ? (a.range.length === 1 ? a.range[0] : E('#VALUE!')) : a;
    };
    const num1 = (f: (x: number) => number): Val => {
      const x = toNum(sc(0));
      if (isErr(x)) return x;
      const r = f(x);
      return Number.isFinite(r) ? r : E('#NUM!');
    };
    const agg = (f: (xs: number[]) => Val): Val => {
      const xs = nums(args);
      return isErr(xs) ? xs : f(xs);
    };
    switch (name) {
      case 'SUM':
        return agg((xs) => xs.reduce((a, b) => a + b, 0));
      case 'PRODUCT':
        return agg((xs) => xs.reduce((a, b) => a * b, 1));
      case 'AVERAGE':
        return agg((xs) => (xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : E('#DIV/0!')));
      case 'MIN':
        return agg((xs) => (xs.length ? Math.min(...xs) : 0));
      case 'MAX':
        return agg((xs) => (xs.length ? Math.max(...xs) : 0));
      case 'COUNT':
        return agg((xs) => xs.length);
      case 'COUNTA': {
        let k = 0;
        for (const a of args) {
          if (isRange(a)) k += a.range.filter((v) => v !== null && v !== '').length;
          else if (a !== null) k++;
        }
        return k;
      }
      case 'ABS':
        return num1(Math.abs);
      case 'SQRT':
        return num1(Math.sqrt);
      case 'INT':
        return num1(Math.floor);
      case 'ROUND': {
        const d = toNum(sc(1));
        if (isErr(d)) return d;
        const f = Math.pow(10, Math.trunc(d));
        return num1((x) => (Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9)) / f);
      }
      case 'POWER': {
        const y = toNum(sc(1));
        if (isErr(y)) return y;
        return num1((x) => Math.pow(x, y));
      }
      case 'MOD': {
        const y = toNum(sc(1));
        if (isErr(y)) return y;
        if (y === 0) return E('#DIV/0!');
        return num1((x) => x - y * Math.floor(x / y));
      }
      case 'AND':
      case 'OR': {
        const bs: boolean[] = [];
        for (const a of args) {
          for (const v of isRange(a) ? a.range : [a]) {
            if (v === null || (isRange(a) && typeof v === 'string')) continue;
            const b = toBool(v);
            if (isErr(b)) return b;
            bs.push(b);
          }
        }
        return name === 'AND' ? bs.every(Boolean) : bs.some(Boolean);
      }
      case 'NOT': {
        const b = toBool(sc(0));
        return isErr(b) ? b : !b;
      }
      case 'LEN': {
        const v = sc(0);
        return isErr(v) ? v : toStr(v).length;
      }
      case 'UPPER':
      case 'LOWER':
      case 'TRIM': {
        const v = sc(0);
        if (isErr(v)) return v;
        const s = toStr(v);
        return name === 'UPPER' ? s.toUpperCase() : name === 'LOWER' ? s.toLowerCase() : s.trim().replace(/\s+/g, ' ');
      }
      case 'LEFT':
      case 'RIGHT': {
        const v = sc(0);
        const k = args.length > 1 ? toNum(sc(1)) : 1;
        if (isErr(v)) return v;
        if (isErr(k)) return k;
        const s = toStr(v);
        return name === 'LEFT' ? s.slice(0, k) : k <= 0 ? '' : s.slice(-k);
      }
      case 'CONCAT':
      case 'CONCATENATE': {
        let s = '';
        for (const a of args) {
          for (const v of isRange(a) ? a.range : [a]) {
            if (isErr(v)) return v;
            s += toStr(v);
          }
        }
        return s;
      }
      default:
        return E('#NAME?');
    }
  }

  return { value };
}

export type Evaluator = ReturnType<typeof createEvaluator>;

// ---------- display ----------

export function display(v: Val, z?: string): string {
  if (v === null) return '';
  if (isErr(v)) return v.err;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (z && z !== 'General') {
      const f = formatNumber(z, v);
      if (f !== null) return f;
    }
    return String(+v.toPrecision(12));
  }
  return v;
}
