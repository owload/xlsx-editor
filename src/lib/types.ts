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
  /** Row heights in px at 100% where they differ from the default (20 px); keys are 0-based rows. */
  rowHeights?: Record<number, number>;
  /**
   * Formatting of a whole column (a column that was filled or colored as a whole): it applies to the cells of the
   * column that do not exist. A cell that exists carries its own complete style; an empty style `{}` on a cell
   * means "deliberately none". Keys are 0-based columns.
   */
  colStyles?: Record<number, Style>;
  /** The same for whole rows; where a row and a column both have a style, the row's wins. */
  rowStyles?: Record<number, Style>;
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
  /** Set when the rectangle is whole columns or whole rows (chosen by their headers), however many cells are shown. */
  whole?: 'col' | 'row';
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
  /** The selection is whole columns or whole rows, chosen by their headers. */
  whole?: 'col' | 'row';
}

export const selRect = (s: Sel): Rect => ({
  r1: Math.min(s.ar, s.fr),
  r2: Math.max(s.ar, s.fr),
  c1: Math.min(s.ac, s.fc),
  c2: Math.max(s.ac, s.fc),
  ...(s.whole && { whole: s.whole }),
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

/**
 * The formatting an empty cell has by inheritance: its row's style, else its column's. A cell that is typed
 * into starts with it, as in a spreadsheet program, so a filled column stays filled.
 */
export const inheritedStyle = (s: Sheet, r: number, c: number): Style | undefined => {
  const st = s.rowStyles?.[r] ?? s.colStyles?.[c];
  return st && Object.keys(st).length ? st : undefined;
};

