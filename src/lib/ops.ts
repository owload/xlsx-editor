import type { Book, Cell, Rect, Style } from './types';
import { getCell, usedSize } from './types';
import { display, NUM_RE, shiftFormula } from './formula';
import type { Evaluator } from './formula';
import { encodeAddr } from './addr';
import { setCells } from './store';

type Edit = { r: number; c: number; cell?: Cell };

const MAX_CELLS = 50_000;

/** An empty cell with no formatting is not stored. */
export const norm = (c: Cell | undefined): Cell | undefined =>
  c && (c.raw !== '' || c.st || c.z) ? c : undefined;

const forRect = (rc: Rect, fn: (r: number, c: number) => void) => {
  if ((rc.r2 - rc.r1 + 1) * (rc.c2 - rc.c1 + 1) > MAX_CELLS) return;
  for (let r = rc.r1; r <= rc.r2; r++) for (let c = rc.c1; c <= rc.c2; c++) fn(r, c);
};

export function patchStyle(book: Book, si: number, rc: Rect, patch: Partial<Style>): Book {
  const s = book.sheets[si];
  const edits: Edit[] = [];
  forRect(rc, (r, c) => {
    const cell = getCell(s, r, c) ?? { raw: '' };
    const st: Record<string, unknown> = { ...cell.st, ...patch };
    for (const k of Object.keys(st)) if (st[k] === undefined || st[k] === false) delete st[k];
    const next: Cell = { ...cell, st: Object.keys(st).length ? (st as Style) : undefined };
    if (!next.st) delete next.st;
    edits.push({ r, c, cell: norm(next) });
  });
  return edits.length ? setCells(book, si, edits) : book;
}

export function setFormat(book: Book, si: number, rc: Rect, z: string | undefined): Book {
  const s = book.sheets[si];
  const edits: Edit[] = [];
  forRect(rc, (r, c) => {
    const cell = getCell(s, r, c) ?? { raw: '' };
    const next: Cell = { ...cell, z };
    if (!z) delete next.z;
    edits.push({ r, c, cell: norm(next) });
  });
  return edits.length ? setCells(book, si, edits) : book;
}

export function changeDecimals(z: string | undefined, d: 1 | -1): string {
  if (!z || z === 'General') return d > 0 ? '0.0' : '0';
  const m = /\.(0+)/.exec(z);
  if (m) {
    const n = Math.max(0, m[1].length + d);
    return z.replace(/\.0+/, n ? '.' + '0'.repeat(n) : '');
  }
  if (d < 0) return z;
  const i = z.search(/[0#](?=[^0#]*$)/);
  return i < 0 ? '0.0' : z.slice(0, i + 1) + '.0' + z.slice(i + 1);
}

export function clearContents(book: Book, si: number, rc: Rect): Book {
  const s = book.sheets[si];
  const edits: Edit[] = [];
  forRect(rc, (r, c) => {
    const cell = getCell(s, r, c);
    if (cell && cell.raw !== '') edits.push({ r, c, cell: norm({ ...cell, raw: '', text: undefined }) });
  });
  return edits.length ? setCells(book, si, edits) : book;
}

export function autosum(book: Book, si: number, r: number, c: number): Book | null {
  const s = book.sheets[si];
  const isNum = (rr: number, cc: number) => {
    const x = getCell(s, rr, cc);
    return !!x && !x.text && x.raw !== '' && (x.raw[0] === '=' || NUM_RE.test(x.raw.trim()));
  };
  let r0 = r;
  while (r0 > 0 && isNum(r0 - 1, c)) r0--;
  if (r0 < r) return setCells(book, si, [{ r, c, cell: { ...(getCell(s, r, c) ?? { raw: '' }), raw: `=SUM(${encodeAddr(r0, c)}:${encodeAddr(r - 1, c)})` } }]);
  let c0 = c;
  while (c0 > 0 && isNum(r, c0 - 1)) c0--;
  if (c0 < c) return setCells(book, si, [{ r, c, cell: { ...(getCell(s, r, c) ?? { raw: '' }), raw: `=SUM(${encodeAddr(r, c0)}:${encodeAddr(r, c - 1)})` } }]);
  return null;
}

/** Sorts rows by column `col`. A header (text above numbers) stays in place. */
export function sortRows(book: Book, si: number, ev: Evaluator, rc: Rect, col: number, asc: boolean): Book {
  const s = book.sheets[si];
  let { r1 } = rc;
  const { r2, c1, c2 } = rc;
  if (r2 <= r1) return book;
  if (typeof ev.value(si, r1, col) === 'string' && typeof ev.value(si, r1 + 1, col) === 'number') r1++;
  const rows: { r: number; key: ReturnType<Evaluator['value']> }[] = [];
  for (let r = r1; r <= r2; r++) rows.push({ r, key: ev.value(si, r, col) });
  const rank = (k: unknown) => (k === null || k === '' ? 3 : typeof k === 'number' ? 0 : typeof k === 'string' ? 1 : 2);
  rows.sort((a, b) => {
    const ra = rank(a.key);
    const rb = rank(b.key);
    if (ra === 3 || rb === 3) return ra === rb ? a.r - b.r : ra === 3 ? 1 : -1;
    let d = 0;
    if (ra !== rb) d = ra - rb;
    else if (typeof a.key === 'number') d = a.key - (b.key as number);
    else d = String(a.key).localeCompare(String(b.key), undefined, { sensitivity: 'base' });
    return (asc ? d : -d) || a.r - b.r;
  });
  const edits: Edit[] = [];
  rows.forEach((row, i) => {
    const dr = r1 + i - row.r;
    for (let c = c1; c <= c2; c++) {
      const cell = getCell(s, row.r, c);
      edits.push({ r: r1 + i, c, cell: cell && cell.raw[0] === '=' && !cell.text ? { ...cell, raw: shiftFormula(cell.raw, dr, 0) } : cell });
    }
  });
  return setCells(book, si, edits);
}

export type Dir = 'down' | 'up' | 'left' | 'right';

/** Fill handle: formulas are shifted, number series are continued. */
export function fillRange(book: Book, si: number, rc: Rect, dir: Dir, count: number): Book {
  const s = book.sheets[si];
  const vertical = dir === 'down' || dir === 'up';
  const fwd = dir === 'down' || dir === 'right';
  const lines = vertical ? rc.c2 - rc.c1 + 1 : rc.r2 - rc.r1 + 1;
  const n = vertical ? rc.r2 - rc.r1 + 1 : rc.c2 - rc.c1 + 1;
  if (lines * (n + count) > MAX_CELLS) return book;
  const edits: Edit[] = [];
  for (let line = 0; line < lines; line++) {
    const at = (i: number) => (vertical ? getCell(s, rc.r1 + i, rc.c1 + line) : getCell(s, rc.r1 + line, rc.c1 + i));
    const src = Array.from({ length: n }, (_, i) => at(i));
    const nums = src.map((x) => (x && !x.text && NUM_RE.test(x.raw.trim()) ? Number(x.raw) : null));
    const linear = n >= 2 && nums.every((x) => x !== null);
    for (let k = 1; k <= count; k++) {
      const idx = fwd ? (k - 1) % n : n - 1 - ((k - 1) % n);
      const pos = fwd ? n - 1 + k : -k;
      const r = vertical ? rc.r1 + pos : rc.r1 + line;
      const c = vertical ? rc.c1 + line : rc.c1 + pos;
      if (r < 0 || c < 0 || r > 1048575 || c > 16383) continue;
      const sc = src[idx];
      if (!sc) {
        edits.push({ r, c });
        continue;
      }
      let cell: Cell;
      if (linear) {
        const a = nums as number[];
        const step = fwd ? a[n - 1] - a[n - 2] : a[1] - a[0];
        const v = fwd ? a[n - 1] + step * k : a[0] - step * k;
        cell = { ...sc, raw: String(+v.toPrecision(15)) };
      } else {
        const d = pos - idx;
        cell = sc.raw[0] === '=' && !sc.text ? { ...sc, raw: shiftFormula(sc.raw, vertical ? d : 0, vertical ? 0 : d) } : sc;
      }
      edits.push({ r, c, cell });
    }
  }
  return edits.length ? setCells(book, si, edits) : book;
}

// ---------- find and replace ----------

const esc = (q: string) => q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function findNext(book: Book, si: number, ev: Evaluator, q: string, matchCase: boolean, from: { r: number; c: number }) {
  if (!q) return null;
  const s = book.sheets[si];
  const { rows, cols } = usedSize(s);
  const total = rows * cols;
  if (!total) return null;
  const needle = matchCase ? q : q.toLowerCase();
  const start = from.r * cols + from.c;
  for (let k = 1; k <= total; k++) {
    const i = (start + k + total * 2) % total;
    const r = Math.floor(i / cols);
    const c = i % cols;
    const cell = s.rows[r]?.[c];
    if (!cell || cell.raw === '') continue;
    const text = display(ev.value(si, r, c), cell.z);
    if ((matchCase ? text : text.toLowerCase()).includes(needle)) return { r, c };
  }
  return null;
}

const replaceText = (raw: string, q: string, to: string, matchCase: boolean) =>
  raw.replace(new RegExp(esc(q), matchCase ? 'g' : 'gi'), () => to);

export function replaceAt(book: Book, si: number, r: number, c: number, q: string, to: string, matchCase: boolean): Book {
  const cell = getCell(book.sheets[si], r, c);
  if (!cell || !q) return book;
  const raw = replaceText(cell.raw, q, to, matchCase);
  return raw === cell.raw ? book : setCells(book, si, [{ r, c, cell: { ...cell, raw } }]);
}

export function replaceAll(book: Book, si: number, q: string, to: string, matchCase: boolean): { book: Book; count: number } {
  if (!q) return { book, count: 0 };
  const s = book.sheets[si];
  const edits: Edit[] = [];
  s.rows.forEach((row, r) =>
    row?.forEach((cell, c) => {
      if (!cell || cell.raw === '') return;
      const raw = replaceText(cell.raw, q, to, matchCase);
      if (raw !== cell.raw) edits.push({ r, c, cell: { ...cell, raw } });
    }),
  );
  return { book: edits.length ? setCells(book, si, edits) : book, count: edits.length };
}
