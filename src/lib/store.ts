import type { Book, Cell, Sheet } from './types';
import { usedSize } from './types';
import { adjustForEdit, renameSheetInFormula } from './formula';

export const emptySheet = (name: string): Sheet => ({ name, rows: [], colWidths: {} });
export const newBook = (): Book => ({ sheets: [emptySheet('Sheet1')], active: 0 });

const mapSheet = (b: Book, i: number, fn: (s: Sheet) => Sheet): Book => ({
  ...b,
  sheets: b.sheets.map((s, k) => (k === i ? fn(s) : s)),
});

export function setCells(b: Book, si: number, edits: { r: number; c: number; cell?: Cell }[]): Book {
  return mapSheet(b, si, (s) => {
    const rows = s.rows.slice();
    const copied = new Set<number>();
    for (const { r, c, cell } of edits) {
      let row = rows[r];
      if (!copied.has(r)) {
        row = (row ?? []).slice();
        rows[r] = row;
        copied.add(r);
      }
      row![c] = cell;
    }
    return { ...s, rows };
  });
}

/** Applies fn to every formula on every sheet. */
function mapFormulas(b: Book, fn: (raw: string, own: string) => string): Book {
  return {
    ...b,
    sheets: b.sheets.map((s) => ({
      ...s,
      rows: s.rows.map((row) =>
        row?.map((cell) => {
          if (!cell || cell.raw[0] !== '=' || cell.text) return cell;
          const raw = fn(cell.raw, s.name);
          return raw === cell.raw ? cell : { ...cell, raw };
        }),
      ),
    })),
  };
}

export function insertLines(b: Book, si: number, axis: 'r' | 'c', at: number, n: number): Book {
  const target = b.sheets[si].name;
  const moved = mapSheet(b, si, (s) => {
    if (axis === 'r') {
      if (at >= s.rows.length) return { ...s, merges: undefined };
      const rows = s.rows.slice();
      rows.splice(at, 0, ...new Array<undefined>(n).fill(undefined));
      return { ...s, rows, merges: undefined };
    }
    const widths: Record<number, number> = {};
    for (const [k, w] of Object.entries(s.colWidths)) widths[Number(k) >= at ? Number(k) + n : Number(k)] = w;
    return {
      ...s,
      colWidths: widths,
      merges: undefined,
      rows: s.rows.map((row) => {
        if (!row || at >= row.length) return row;
        const r = row.slice();
        r.splice(at, 0, ...new Array<undefined>(n).fill(undefined));
        return r;
      }),
    };
  });
  return mapFormulas(moved, (raw, own) => adjustForEdit(raw, own, target, axis, at, n));
}

export function deleteLines(b: Book, si: number, axis: 'r' | 'c', at: number, n: number): Book {
  const target = b.sheets[si].name;
  const moved = mapSheet(b, si, (s) => {
    if (axis === 'r') {
      const rows = s.rows.slice();
      rows.splice(at, n);
      return { ...s, rows, merges: undefined };
    }
    const widths: Record<number, number> = {};
    for (const [k, w] of Object.entries(s.colWidths)) {
      const c = Number(k);
      if (c < at) widths[c] = w;
      else if (c >= at + n) widths[c - n] = w;
    }
    return {
      ...s,
      colWidths: widths,
      merges: undefined,
      rows: s.rows.map((row) => {
        if (!row) return row;
        const r = row.slice();
        r.splice(at, n);
        return r;
      }),
    };
  });
  return mapFormulas(moved, (raw, own) => adjustForEdit(raw, own, target, axis, at, -n));
}

const BAD_NAME = /[[\]:*?/\\]/g;

export function uniqueName(b: Book, wanted: string, except = -1): string {
  const base = (wanted.replace(BAD_NAME, '').trim() || 'Sheet').slice(0, 31);
  const taken = (n: string) => b.sheets.some((s, i) => i !== except && s.name.toLowerCase() === n.toLowerCase());
  let name = base;
  for (let k = 2; taken(name); k++) name = base.slice(0, 31 - String(k).length - 1) + ' ' + k;
  return name;
}

export function addSheet(b: Book): Book {
  const name = uniqueName(b, `Sheet${b.sheets.length + 1}`);
  return { sheets: [...b.sheets, emptySheet(name)], active: b.sheets.length };
}

export function renameSheet(b: Book, si: number, wanted: string): Book {
  const from = b.sheets[si].name;
  const to = uniqueName(b, wanted, si);
  if (to === from) return b;
  const renamed = mapSheet(b, si, (s) => ({ ...s, name: to }));
  return mapFormulas(renamed, (raw) => renameSheetInFormula(raw, from, to));
}

export function moveSheet(b: Book, from: number, to: number): Book {
  if (from === to || from < 0 || to < 0 || from >= b.sheets.length || to >= b.sheets.length) return b;
  const sheets = b.sheets.slice();
  const [moved] = sheets.splice(from, 1);
  sheets.splice(to, 0, moved);
  return { sheets, active: sheets.indexOf(b.sheets[b.active]) };
}

export function deleteSheet(b: Book, si: number): Book {
  if (b.sheets.length < 2) return b;
  const sheets = b.sheets.filter((_, i) => i !== si);
  return { sheets, active: Math.min(b.active > si ? b.active - 1 : b.active, sheets.length - 1) };
}

export { usedSize };

// ---------- history ----------

export interface History {
  book: Book;
  past: Book[];
  future: Book[];
  /** Bumped by every content change (edit, undo, redo); `savedRev` is the rev that was last saved. */
  rev: number;
  savedRev: number;
}

export type Action =
  | { type: 'commit'; fn: (b: Book) => Book }
  | { type: 'silent'; fn: (b: Book) => Book }
  | { type: 'undo' }
  | { type: 'redo' }
  | { type: 'load'; book: Book }
  | { type: 'saved'; rev: number };

const MAX_HISTORY = 100;

export const initialHistory = (book: Book): History => ({ book, past: [], future: [], rev: 0, savedRev: 0 });
export const isDirty = (h: History) => h.rev !== h.savedRev;

export function reducer(h: History, a: Action): History {
  switch (a.type) {
    case 'commit': {
      const book = a.fn(h.book);
      if (book === h.book) return h;
      return { ...h, book, past: [...h.past, h.book].slice(-MAX_HISTORY), future: [], rev: h.rev + 1 };
    }
    case 'silent': {
      const book = a.fn(h.book);
      return book === h.book ? h : { ...h, book };
    }
    case 'undo': {
      const prev = h.past[h.past.length - 1];
      if (!prev) return h;
      return { ...h, book: { ...prev, active: h.book.active }, past: h.past.slice(0, -1), future: [h.book, ...h.future], rev: h.rev + 1 };
    }
    case 'redo': {
      const next = h.future[0];
      if (!next) return h;
      return { ...h, book: { ...next, active: h.book.active }, past: [...h.past, h.book], future: h.future.slice(1), rev: h.rev + 1 };
    }
    case 'load':
      return initialHistory(a.book);
    case 'saved':
      return { ...h, savedRev: a.rev };
  }
}
