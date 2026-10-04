import { describe, expect, test } from 'vitest';
import type { CanvasFactory, DrawContext } from '../lib/canvas';
import { MAX_PREVIEW_SOURCE_BYTES, layoutSheetPreview, renderSheetPreview } from '../lib/preview';
import type { Book } from '../lib/types';
import { writeWorkbook } from '../lib/xlsx-io';

const book = (rows: Book['sheets'][0]['rows'], colWidths: Record<number, number> = {}): Book => ({
  sheets: [{ name: 'S', rows, colWidths }],
  active: 0,
});

function fakeCanvas() {
  const calls: string[] = [];
  const g: DrawContext = {
    fillStyle: '',
    font: '',
    textBaseline: '',
    textAlign: '',
    fillRect: (x, y, w, h) => calls.push(`rect ${x},${y},${w},${h} ${String(g.fillStyle)}`),
    fillText: (t, x, y) => calls.push(`text ${t} @${x},${y} ${g.textAlign} ${g.font.split(' ')[0]} ${String(g.fillStyle)}`),
    scale: (x) => calls.push(`scale ${x}`),
  };
  const sizes: [number, number][] = [];
  const factory: CanvasFactory = (w, h) => {
    sizes.push([w, h]);
    return { context: g, toPng: async () => new Uint8Array([137, 80, 78, 71]) };
  };
  return { calls, sizes, factory };
}

describe('layoutSheetPreview', () => {
  test('lays out a landscape picture whose width is the requested size', () => {
    const layout = layoutSheetPreview(book([[{ raw: 'a' }]]), 0, 360)!;
    expect([layout.width, layout.height]).toEqual([360, 270]);
    expect(layout.scale).toBeCloseTo(0.6);
  });

  test('shows column letters from A and row numbers from 1, as many as fit', () => {
    const layout = layoutSheetPreview(book([[{ raw: 'a' }]]), 0, 360)!;
    expect(layout.columns.slice(0, 3).map((c) => c.label)).toEqual(['A', 'B', 'C']);
    expect(layout.columns.length).toBe(8); // (600 - 40) / 72, rounded up
    expect(layout.rows[0].label).toBe('1');
    expect(layout.rows.length).toBe(21); // whole rows only
  });

  test('uses the column widths of the sheet', () => {
    const layout = layoutSheetPreview(book([[{ raw: 'a' }]], { 0: 200 }), 0, 360)!;
    expect(layout.columns[0].w).toBe(200);
    expect(layout.columns[1].x).toBe(40 + 200);
  });

  test('puts text left, numbers right, errors and booleans centered, unless the cell says otherwise', () => {
    const layout = layoutSheetPreview(
      book([[{ raw: 'text' }, { raw: '42' }, { raw: '=1/0' }, { raw: 'TRUE' }, { raw: 'x', st: { h: 'center' } }]]),
      0,
      360,
    )!;
    expect(layout.cells.map((c) => [c.text, c.align])).toEqual([
      ['text', 'left'],
      ['42', 'right'],
      ['#DIV/0!', 'center'],
      ['TRUE', 'center'],
      ['x', 'center'],
    ]);
    expect(layout.cells[2].error).toBe(true);
    expect(layout.cells[2].color).toBe('#d1242f');
  });

  test('shows the evaluated value of a formula and applies the number format', () => {
    const layout = layoutSheetPreview(
      book([[{ raw: '2' }, { raw: '=A1*3' }, { raw: '0.5', z: '0%' }]]),
      0,
      360,
    )!;
    expect(layout.cells.map((c) => c.text)).toEqual(['2', '6', '50%']);
  });

  test('keeps bold, italic, size and colors; skips empty cells', () => {
    const layout = layoutSheetPreview(
      book([[{ raw: 'h', st: { b: true, i: true, sz: 22, color: '#112233', bg: '#ffee00' } }, undefined, { raw: '' }]]),
      0,
      360,
    )!;
    expect(layout.cells).toHaveLength(1);
    expect(layout.cells[0]).toMatchObject({ bold: true, italic: true, fontPx: 28, color: '#112233', bg: '#ffee00' });
  });

  test('cuts text that does not fit the column with an ellipsis', () => {
    const layout = layoutSheetPreview(book([[{ raw: 'a very long piece of text that cannot fit' }]]), 0, 360)!;
    expect(layout.cells[0].text.endsWith('…')).toBe(true);
    expect(layout.cells[0].text.length).toBeLessThan(20);
  });

  test('leaves out cells outside the visible corner', () => {
    const far: Book['sheets'][0]['rows'] = [];
    far[100] = [];
    far[100][40] = { raw: 'far away' };
    const layout = layoutSheetPreview(book(far), 0, 360)!;
    expect(layout.cells).toEqual([]);
  });

  test('has nothing to show for an empty sheet or a missing one', () => {
    expect(layoutSheetPreview(book([]), 0, 360)).toBeNull();
    expect(layoutSheetPreview(book([[]]), 0, 360)).toBeNull();
    expect(layoutSheetPreview(book([[{ raw: 'a' }]]), 3, 360)).toBeNull();
  });
});

describe('renderSheetPreview', () => {
  test('draws a file the editor wrote: white page, headers, grid, the values, and returns the canvas PNG', async () => {
    const bytes = await writeWorkbook(book([[{ raw: 'Name', st: { b: true } }, { raw: '=2+3' }]]));
    const c = fakeCanvas();
    const png = await renderSheetPreview(bytes, 360, c.factory);
    expect([...png!]).toEqual([137, 80, 78, 71]);
    expect(c.sizes).toEqual([[360, 270]]);
    expect(c.calls[0]).toBe('rect 0,0,360,270 #ffffff');
    expect(c.calls).toContain('scale 0.6');
    expect(c.calls.some((x) => x.startsWith('text A @'))).toBe(true);
    expect(c.calls.some((x) => x.startsWith('text 1 @'))).toBe(true);
    expect(c.calls.some((x) => x.startsWith('text Name @') && x.includes('bold'))).toBe(true);
    expect(c.calls.some((x) => x.startsWith('text 5 @') && x.includes(' right '))).toBe(true);
  });

  test('has no preview for empty input, a file over the limit, or an empty workbook', async () => {
    const c = fakeCanvas();
    expect(await renderSheetPreview(new Uint8Array(), 360, c.factory)).toBeNull();
    expect(await renderSheetPreview(new Uint8Array(MAX_PREVIEW_SOURCE_BYTES + 1), 360, c.factory)).toBeNull();
    expect(await renderSheetPreview(await writeWorkbook(book([])), 360, c.factory)).toBeNull();
    expect(c.sizes).toEqual([]);
  });

  test('fails for a file that is not a workbook, instead of drawing something', async () => {
    await expect(renderSheetPreview(new TextEncoder().encode('not a workbook'), 360, fakeCanvas().factory)).rejects.toThrow();
  });

  test('has no preview where there is no canvas', async () => {
    const bytes = await writeWorkbook(book([[{ raw: 'a' }]]));
    expect(await renderSheetPreview(bytes, 360, () => null)).toBeNull();
    expect(await renderSheetPreview(bytes, 360)).toBeNull(); // happy-dom has no OffscreenCanvas
  });

  test('does not change the bytes it was given', async () => {
    const bytes = await writeWorkbook(book([[{ raw: 'a' }]]));
    const copy = new Uint8Array(bytes);
    await renderSheetPreview(bytes, 360, fakeCanvas().factory);
    expect([...bytes]).toEqual([...copy]);
  });
});
