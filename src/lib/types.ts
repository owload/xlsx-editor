export interface Range {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

export interface Style {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  /** Font size in pt (default 11). */
  sz?: number;
  /** Text color, #rrggbb only. */
  color?: string;
  /** Fill color, #rrggbb only. */
  bg?: string;
  h?: 'left' | 'center' | 'right';
  /** Thin borders on all sides. */
  bd?: boolean;
}

export const isColor = (c: unknown): c is string => typeof c === 'string' && /^#[0-9a-fA-F]{6}$/.test(c);

export interface Cell {
  /** What the user sees in the formula bar: a number, text or "=formula". */
  raw: string;
  /** Excel number format (e.g. "0.00%" or "yyyy-mm-dd"). */
  z?: string;
  /** The value is text even if it looks like a number. */
  text?: boolean;
  st?: Style;
}

export interface Sheet {
  name: string;
  rows: ((Cell | undefined)[] | undefined)[];
  colWidths: Record<number, number>;
  merges?: Range[];
}

export interface Book {
  sheets: Sheet[];
  active: number;
}

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface Sel {
  /** Active cell (anchor). */
  ar: number;
  ac: number;
  /** Moving end of the selection. */
  fr: number;
  fc: number;
  /** Do not scroll to the cell (whole row/column selection). */
  noScroll?: boolean;
  /** Previously selected ranges (Ctrl+click). */
  extra?: Rect[];
}

export const selRect = (s: Sel) => ({
  r1: Math.min(s.ar, s.fr),
  r2: Math.max(s.ar, s.fr),
  c1: Math.min(s.ac, s.fc),
  c2: Math.max(s.ac, s.fc),
});

export const getCell = (s: Sheet, r: number, c: number): Cell | undefined => s.rows[r]?.[c];

export function usedSize(s: Sheet) {
  let cols = 0;
  for (const row of s.rows) if (row && row.length > cols) cols = row.length;
  return { rows: s.rows.length, cols };
}

/** All selected ranges: the previous ones and the current one. */
export const selRects = (s: Sel): Rect[] => [...(s.extra ?? []), selRect(s)];

/** Merged non-overlapping row or column intervals, in descending order. */
export function intervals(rects: Rect[], axis: 'r' | 'c'): [number, number][] {
  const list = rects.map((x): [number, number] => (axis === 'r' ? [x.r1, x.r2] : [x.c1, x.c2])).sort((a, b) => a[0] - b[0]);
  const out: [number, number][] = [];
  for (const [a, b] of list) {
    const last = out[out.length - 1];
    if (last && a <= last[1] + 1) last[1] = Math.max(last[1], b);
    else out.push([a, b]);
  }
  return out.reverse();
}
