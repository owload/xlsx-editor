import { encodeCol } from './addr';
import { browserCanvas, type CanvasFactory } from './canvas';
import { createEvaluator, display, isErr } from './formula';
import type { Book, Sheet } from './types';
import { usedSize } from './types';
import { readWorkbook } from './xlsx-io';

/**
 * The preview of a spreadsheet for the file grid (owload-docs/decisions/0020): the top-left corner of the
 * active sheet drawn as a grid with column letters, row numbers and the displayed values. Reads the
 * file and changes nothing.
 */

/** Bigger files are not parsed just for a thumbnail: parsing runs on the page and cannot be interrupted. */
export const MAX_PREVIEW_SOURCE_BYTES = 20 * 1024 * 1024;

// The sizes of the editor at 100%: 20 px rows, 72 px columns; a corner of this size is shown.
const ROW_H = 20;
const COL_W = 72;
const HEAD_W = 40;
const HEAD_H = 20;
const VIEW_W = 600;
const VIEW_H = 450;
const FONT_PX = 14; // 11 pt in the editor
const CHAR_W = 0.55; // average glyph width relative to the font size

export interface PreviewCell {
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  align: 'left' | 'center' | 'right';
  bold: boolean;
  italic: boolean;
  fontPx: number;
  color?: string;
  bg?: string;
  error: boolean;
}

export interface SheetPreviewLayout {
  /** The size of the PNG. */
  width: number;
  height: number;
  /** PNG pixels per layout unit. */
  scale: number;
  columns: { label: string; x: number; w: number }[];
  rows: { label: string; y: number }[];
  cells: PreviewCell[];
}

const fit = (text: string, widthPx: number, fontPx: number) => {
  const max = Math.max(1, Math.floor((widthPx - 8) / (fontPx * CHAR_W)));
  return text.length > max ? text.slice(0, Math.max(1, max - 1)) + '…' : text;
};

/** Positions the visible corner of the sheet `si` of `book`. Null if the sheet has no data. */
export function layoutSheetPreview(book: Book, si: number, size: number): SheetPreviewLayout | null {
  const sheet: Sheet | undefined = book.sheets[si];
  if (!sheet) return null;
  const used = usedSize(sheet);
  if (used.rows === 0 || used.cols === 0) return null;

  const width = Math.max(32, Math.min(Math.round(size), 4096));
  const height = Math.round(width * 0.75);
  const scale = width / VIEW_W;
  const ev = createEvaluator(book);

  const columns: SheetPreviewLayout['columns'] = [];
  for (let c = 0, x = HEAD_W; x < VIEW_W && c < 100; c++) {
    const w = sheet.colWidths[c] ?? COL_W;
    columns.push({ label: encodeCol(c), x, w });
    x += w;
  }
  const rows: SheetPreviewLayout['rows'] = [];
  for (let r = 0, y = HEAD_H; y + ROW_H <= VIEW_H && r < 100; r++, y += ROW_H) rows.push({ label: String(r + 1), y });

  const cells: PreviewCell[] = [];
  for (const row of rows.keys()) {
    for (const [c, col] of columns.entries()) {
      const cell = sheet.rows[row]?.[c];
      if (!cell || (cell.raw === '' && !cell.st)) continue;
      const v = ev.value(si, row, c);
      const st = cell.st;
      const fontPx = st?.sz ? (st.sz * FONT_PX) / 11 : FONT_PX;
      const error = isErr(v);
      cells.push({
        x: col.x,
        y: rows[row].y,
        w: col.w,
        h: ROW_H,
        text: fit(display(v, cell.z), col.w, fontPx),
        align: st?.h ?? (typeof v === 'number' ? 'right' : error || typeof v === 'boolean' ? 'center' : 'left'),
        bold: !!st?.b,
        italic: !!st?.i,
        fontPx,
        color: error ? '#d1242f' : st?.color,
        bg: st?.bg,
        error,
      });
    }
  }
  return { width, height, scale, columns, rows, cells };
}

const FONT = 'Calibri, Carlito, "Segoe UI", system-ui, sans-serif';

/** The PNG of the top-left corner of the active sheet, at most `size` pixels wide; null if there is nothing to show. */
export async function renderSheetPreview(
  data: Uint8Array,
  size: number,
  makeCanvas: CanvasFactory = browserCanvas,
): Promise<Uint8Array | null> {
  if (data.byteLength === 0 || data.byteLength > MAX_PREVIEW_SOURCE_BYTES) return null;
  const book = await readWorkbook(data);
  const layout = layoutSheetPreview(book, book.active, size);
  if (!layout) return null;
  const canvas = makeCanvas(layout.width, layout.height);
  if (!canvas) return null;

  const g = canvas.context;
  const { columns, rows, cells } = layout;
  g.fillStyle = '#ffffff';
  g.fillRect(0, 0, layout.width, layout.height);
  g.scale(layout.scale, layout.scale);

  // Headers
  g.fillStyle = '#f6f8fa';
  g.fillRect(0, 0, VIEW_W, HEAD_H);
  g.fillRect(0, 0, HEAD_W, VIEW_H);
  // Cell fills
  for (const cell of cells) {
    if (!cell.bg) continue;
    g.fillStyle = cell.bg;
    g.fillRect(cell.x, cell.y, cell.w, cell.h);
  }
  // Grid lines
  g.fillStyle = '#d8dee4';
  for (const col of columns) g.fillRect(col.x, 0, 1, VIEW_H);
  for (const row of rows) g.fillRect(0, row.y, VIEW_W, 1);
  g.fillRect(0, HEAD_H, VIEW_W, 1);
  g.fillRect(HEAD_W, 0, 1, VIEW_H);

  g.textBaseline = 'middle';
  g.fillStyle = '#656d76';
  g.font = `${FONT_PX - 1}px ${FONT}`;
  g.textAlign = 'center';
  for (const col of columns) g.fillText(col.label, col.x + col.w / 2, HEAD_H / 2);
  for (const row of rows) g.fillText(row.label, HEAD_W / 2, row.y + ROW_H / 2);

  for (const cell of cells) {
    g.fillStyle = cell.color ?? '#1f2328';
    g.font = `${cell.italic ? 'italic ' : ''}${cell.bold ? 'bold ' : ''}${cell.fontPx}px ${FONT}`;
    g.textAlign = cell.align;
    const x = cell.align === 'left' ? cell.x + 4 : cell.align === 'right' ? cell.x + cell.w - 4 : cell.x + cell.w / 2;
    g.fillText(cell.text, x, cell.y + cell.h / 2);
  }
  return canvas.toPng();
}
