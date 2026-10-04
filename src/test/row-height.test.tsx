import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, test, vi } from 'vitest';
import { XlsxEditor, type XlsxEditorHandle } from '../index';
import { layoutSheetPreview } from '../lib/preview';
import {
  DEFAULT_ROW_HEIGHT,
  MAX_ROW_HEIGHT,
  MIN_ROW_HEIGHT,
  deleteLines,
  initialHistory,
  insertLines,
  isDirty,
  reducer,
  setColWidth,
  setRowHeight,
  shiftRowHeightsForDelete,
  shiftRowHeightsForInsert,
} from '../lib/store';
import type { Book } from '../lib/types';
import { readWorkbook, writeWorkbook } from '../lib/xlsx-io';
import { unzip, zip } from '../lib/zip';

(globalThis as unknown as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true;

const book = (rowHeights?: Record<number, number>): Book => ({
  sheets: [{ name: 'S', rows: [[{ raw: 'a' }], [{ raw: 'b' }], [{ raw: 'c' }]], colWidths: {}, rowHeights }],
  active: 0,
});
const heights = (b: Book) => b.sheets[0].rowHeights;

describe('setRowHeight', () => {
  test('sets a height, replaces it, and the default removes the entry', () => {
    const b = setRowHeight(book(), 0, 1, 50);
    expect(heights(b)).toEqual({ 1: 50 });
    expect(heights(setRowHeight(b, 0, 1, 70))).toEqual({ 1: 70 });
    expect(heights(setRowHeight(b, 0, 1, DEFAULT_ROW_HEIGHT))).toEqual({});
    expect(heights(setRowHeight(b, 0, 1, undefined))).toEqual({});
  });

  test('keeps a row within what a row can be', () => {
    expect(heights(setRowHeight(book(), 0, 0, 1))).toEqual({ 0: MIN_ROW_HEIGHT });
    expect(heights(setRowHeight(book(), 0, 0, 5000))).toEqual({ 0: MAX_ROW_HEIGHT });
    expect(heights(setRowHeight(book(), 0, 0, 33.6))).toEqual({ 0: 34 });
  });

  test('does not change the book it was given', () => {
    const b = book({ 0: 30 });
    setRowHeight(b, 0, 0, 90);
    expect(heights(b)).toEqual({ 0: 30 });
  });
});

describe('setColWidth', () => {
  test('sets, clamps to a minimum, and restores the default', () => {
    const b = setColWidth(book(), 0, 2, 150);
    expect(b.sheets[0].colWidths).toEqual({ 2: 150 });
    expect(setColWidth(b, 0, 2, 1).sheets[0].colWidths).toEqual({ 2: 12 });
    expect(setColWidth(b, 0, 2, undefined).sheets[0].colWidths).toEqual({});
  });
});

describe('row heights when rows are inserted or deleted', () => {
  test('rows from the insertion on move down, with their heights', () => {
    expect(shiftRowHeightsForInsert({ 0: 30, 2: 40, 5: 50 }, 2, 3)).toEqual({ 0: 30, 5: 40, 8: 50 });
    expect(shiftRowHeightsForInsert(undefined, 2, 3)).toBeUndefined();
  });

  test('the deleted rows lose their heights and the rows below move up', () => {
    expect(shiftRowHeightsForDelete({ 0: 30, 2: 40, 3: 45, 5: 50 }, 2, 2)).toEqual({ 0: 30, 3: 50 });
  });

  test('applies to the sheet, also for rows beyond the last one with data', () => {
    const inserted = insertLines(book({ 1: 50, 10: 60 }), 0, 'r', 1, 2);
    expect(heights(inserted)).toEqual({ 3: 50, 12: 60 });
    expect(heights(deleteLines(book({ 0: 30, 1: 50, 2: 60 }), 0, 'r', 1, 1))).toEqual({ 0: 30, 1: 60 });
  });

  test('inserting columns does not touch row heights', () => {
    expect(heights(insertLines(book({ 1: 50 }), 0, 'c', 0, 1))).toEqual({ 1: 50 });
  });
});

describe('a change of size is a step of the history', () => {
  test('it makes the document changed, and can be undone and redone', () => {
    let h = initialHistory(book());
    h = reducer(h, { type: 'commit', fn: (b) => setRowHeight(b, 0, 0, 80) });
    expect(isDirty(h)).toBe(true);
    expect(heights(h.book)).toEqual({ 0: 80 });
    h = reducer(h, { type: 'undo' });
    expect(heights(h.book) ?? {}).toEqual({}); // undo is a change of the document too, as for any other edit
    h = reducer(h, { type: 'redo' });
    expect(heights(h.book)).toEqual({ 0: 80 });
  });
});

describe('row heights in the file', () => {
  test('are written in points and read back as the same pixels, also for rows without cells', async () => {
    const b: Book = { sheets: [{ name: 'S', rows: [[{ raw: 'a' }]], colWidths: {}, rowHeights: { 0: 40, 4: 100 } }], active: 0 };
    const bytes = await writeWorkbook(b);
    const xml = new TextDecoder().decode((await unzip(bytes)).get('xl/worksheets/sheet1.xml'));
    expect(xml).toContain('<row r="1" ht="30.00" customHeight="1">');
    expect(xml).toContain('<row r="5" ht="75.00" customHeight="1"></row>');
    const read = await readWorkbook(bytes);
    expect(read.sheets[0].rowHeights).toEqual({ 0: 40, 4: 100 });
  });

  test('a sheet without custom heights writes no height attributes', async () => {
    const xml = new TextDecoder().decode((await unzip(await writeWorkbook(book()))).get('xl/worksheets/sheet1.xml'));
    expect(xml).not.toContain('ht=');
  });

  test('are read from a file made by another program, keeping the default and clamping silly values', async () => {
    const files = await unzip(await writeWorkbook(book()));
    const main = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    files.set(
      'xl/worksheets/sheet1.xml',
      new TextEncoder().encode(
        `<worksheet ${main}><sheetData><row r="1" ht="15" customHeight="1"><c r="A1"><v>1</v></c></row><row r="2" ht="30"/><row r="3" ht="9999" customHeight="1"/><row r="4" ht="0"/></sheetData></worksheet>`,
      ),
    );
    const read = await readWorkbook(await zip(files));
    expect(read.sheets[0].rowHeights).toEqual({ 1: 40, 2: MAX_ROW_HEIGHT }); // 15 pt is the default, 0 is nothing
  });
});

describe('row heights in the preview', () => {
  test('a tall row takes its height, so fewer rows fit and the cells below move down', () => {
    const plain = layoutSheetPreview(book(), 0, 360)!;
    const tall = layoutSheetPreview(book({ 0: 100 }), 0, 360)!;
    expect(tall.rows[0].h).toBe(100);
    expect(tall.rows[1].y).toBe(plain.rows[1].y + 80);
    expect(tall.rows.length).toBeLessThan(plain.rows.length);
    expect(tall.cells.find((c) => c.text === 'a')!.h).toBe(100);
  });
});

describe('resizing in the editor', () => {
  async function open() {
    const bytes = await writeWorkbook(book());
    const container = document.createElement('div');
    document.body.appendChild(container);
    const root = createRoot(container);
    const ref = createRef<XlsxEditorHandle>();
    const onSave = vi.fn();
    const onDirty = vi.fn();
    await act(async () => { root.render(createElement(XlsxEditor, { data: bytes, fileName: 'a.xlsx', onSave, onDirtyChange: onDirty, ref })); });
    await act(async () => { await new Promise((r) => setTimeout(r, 80)); });
    const mouse = (type: string, y: number, x = 0) => act(async () => { window.dispatchEvent(new MouseEvent(type, { clientX: x, clientY: y, bubbles: true })); });
    const dragRow = async (index: number, dy: number) => {
      const grip = container.querySelectorAll<HTMLElement>('.xe-rowheads .xe-rgrip')[index];
      await act(async () => { grip.dispatchEvent(new MouseEvent('mousedown', { clientX: 5, clientY: 100, bubbles: true, cancelable: true })); });
      await mouse('mousemove', 100 + dy);
      await mouse('mouseup', 100 + dy);
    };
    const saved = async () => {
      await act(async () => { await ref.current!.save(); });
      return readWorkbook(onSave.mock.calls.at(-1)![0] as Uint8Array);
    };
    const headerHeight = (i: number) => container.querySelectorAll<HTMLElement>('.xe-rowheads .xe-hd')[i].style.height;
    return { container, ref, onDirty, dragRow, mouse, saved, headerHeight, unmount: () => act(async () => root.unmount()) };
  }

  test('dragging the edge of a row header changes the height of that row, and only that row', async () => {
    const e = await open();
    expect(e.headerHeight(1)).toBe('20px');
    await e.dragRow(1, 30);
    expect(e.headerHeight(1)).toBe('50px');
    expect(e.headerHeight(0)).toBe('20px');
    expect(e.headerHeight(2)).toBe('20px');
    expect((await e.saved()).sheets[0].rowHeights).toEqual({ 1: 50 });
    await e.unmount();
  });

  test('shows the new height while dragging, but changes the document only when the mouse is released', async () => {
    const e = await open();
    const grip = e.container.querySelectorAll<HTMLElement>('.xe-rowheads .xe-rgrip')[0];
    await act(async () => { grip.dispatchEvent(new MouseEvent('mousedown', { clientX: 5, clientY: 100, bubbles: true, cancelable: true })); });
    await e.mouse('mousemove', 140);
    expect(e.headerHeight(0)).toBe('60px');
    expect(e.ref.current!.isDirty()).toBe(false);
    await e.mouse('mouseup', 140);
    expect(e.ref.current!.isDirty()).toBe(true);
    await e.unmount();
  });

  test('cannot make a row smaller than the minimum or bigger than the maximum', async () => {
    const e = await open();
    await e.dragRow(0, -500);
    expect(e.headerHeight(0)).toBe(`${MIN_ROW_HEIGHT}px`);
    await e.dragRow(0, 5000);
    expect(e.headerHeight(0)).toBe(`${MAX_ROW_HEIGHT}px`);
    await e.unmount();
  });

  test('a click on the edge that moves nothing does not change the document', async () => {
    const e = await open();
    await e.dragRow(0, 0);
    expect(e.ref.current!.isDirty()).toBe(false);
    await e.unmount();
  });

  test('double-clicking the edge restores the default height', async () => {
    const e = await open();
    await e.dragRow(2, 40);
    expect(e.headerHeight(2)).toBe('60px');
    const grip = e.container.querySelectorAll<HTMLElement>('.xe-rowheads .xe-rgrip')[2];
    await act(async () => { grip.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true })); });
    expect(e.headerHeight(2)).toBe('20px');
    expect((await e.saved()).sheets[0].rowHeights ?? {}).toEqual({});
    await e.unmount();
  });

  test('the text of a tall row is centered in it, and the row below starts under it', async () => {
    const e = await open();
    await e.dragRow(0, 40);
    const cells = [...e.container.querySelectorAll<HTMLElement>('.xe-cell')];
    const a = cells.find((c) => c.textContent === 'a')!;
    const b = cells.find((c) => c.textContent === 'b')!;
    expect(a.style.height).toBe('60px');
    expect(a.style.lineHeight).toBe('60px');
    expect(b.style.top).toBe('60px');
    await e.unmount();
  });

  test('resizing a column is a step of the history too, so the document becomes changed', async () => {
    const e = await open();
    const grip = e.container.querySelectorAll<HTMLElement>('.xe-colheads .xe-grip')[0];
    await act(async () => { grip.dispatchEvent(new MouseEvent('mousedown', { clientX: 100, clientY: 5, bubbles: true, cancelable: true })); });
    await e.mouse('mousemove', 5, 160);
    await e.mouse('mouseup', 5, 160);
    expect(e.ref.current!.isDirty()).toBe(true);
    expect((await e.saved()).sheets[0].colWidths[0]).toBe(132);
    await e.unmount();
  });
});
