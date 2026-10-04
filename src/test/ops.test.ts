import { describe, expect, test } from 'vitest';
import { createEvaluator } from '../lib/formula';
import { autosum, changeDecimals, fillRange, patchStyle, replaceAll, sortRows } from '../lib/ops';
import { deleteLines, insertLines, moveSheet, newBook, reducer, initialHistory, isDirty, renameSheet, setCells } from '../lib/store';
import type { Book } from '../lib/types';

const fill = (cells: [number, number, string][]): Book =>
  setCells(newBook(), 0, cells.map(([r, c, raw]) => ({ r, c, cell: { raw } })));
const raw = (b: Book, r: number, c: number) => b.sheets[0].rows[r]?.[c]?.raw;

describe('fill handle', () => {
  test('continues a numeric series', () => {
    const b = fillRange(fill([[0, 0, '1'], [1, 0, '3']]), 0, { r1: 0, c1: 0, r2: 1, c2: 0 }, 'down', 2);
    expect([2, 3].map((r) => raw(b, r, 0))).toEqual(['5', '7']);
  });

  test('shifts relative formula references', () => {
    const b = fillRange(fill([[0, 0, '1'], [0, 1, '=A1*2']]), 0, { r1: 0, c1: 1, r2: 0, c2: 1 }, 'down', 2);
    expect([1, 2].map((r) => raw(b, r, 1))).toEqual(['=A2*2', '=A3*2']);
  });
});

describe('sorting', () => {
  test('sorts rows by a column and keeps a text header in place', () => {
    const b0 = fill([[0, 0, 'n'], [1, 0, '3'], [2, 0, '1'], [3, 0, '2']]);
    const b = sortRows(b0, 0, createEvaluator(b0), { r1: 0, c1: 0, r2: 3, c2: 0 }, 0, true);
    expect([0, 1, 2, 3].map((r) => raw(b, r, 0))).toEqual(['n', '1', '2', '3']);
  });
});

describe('structure edits', () => {
  test('insert and delete rows keep formulas consistent', () => {
    const b0 = fill([[0, 0, '1'], [1, 0, '2'], [2, 0, '=SUM(A1:A2)']]);
    const ins = insertLines(b0, 0, 'r', 1, 1);
    expect(raw(ins, 3, 0)).toBe('=SUM(A1:A3)');
    const del = deleteLines(b0, 0, 'r', 0, 1);
    expect(raw(del, 1, 0)).toBe('=SUM(A1:A1)');
  });

  test('renaming a sheet rewrites references to it', () => {
    let b = fill([[0, 0, '=Sheet1!B1']]);
    b = renameSheet(b, 0, 'Totals');
    expect(raw(b, 0, 0)).toBe('=Totals!B1');
  });

  test('sheet names are made unique and sanitized', () => {
    let b: Book = { sheets: [...newBook().sheets, { name: 'B', rows: [], colWidths: {} }], active: 0 };
    b = renameSheet(b, 1, 'Sheet1');
    expect(b.sheets[1].name).toBe('Sheet1 2');
    b = renameSheet(b, 1, 'a/b:c');
    expect(b.sheets[1].name).toBe('abc');
  });

  test('moving a sheet keeps the same sheet active', () => {
    let b: Book = { sheets: ['A', 'B', 'C'].map((name) => ({ name, rows: [], colWidths: {} })), active: 2 };
    b = moveSheet(b, 0, 2);
    expect(b.sheets.map((s) => s.name)).toEqual(['B', 'C', 'A']);
    expect(b.sheets[b.active].name).toBe('C');
  });
});

describe('formatting', () => {
  test('patchStyle toggles flags and drops empty styles', () => {
    let b = patchStyle(newBook(), 0, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: true, bg: '#ffff00' });
    expect(b.sheets[0].rows[0]![0]!.st).toEqual({ b: true, bg: '#ffff00' });
    b = patchStyle(b, 0, { r1: 0, c1: 0, r2: 0, c2: 0 }, { b: false, bg: undefined });
    expect(b.sheets[0].rows[0]![0]).toBeUndefined();
  });

  test('changeDecimals', () => {
    expect(changeDecimals(undefined, 1)).toBe('0.0');
    expect(changeDecimals('0.00', 1)).toBe('0.000');
    expect(changeDecimals('0.00', -1)).toBe('0.0');
    expect(changeDecimals('0.0', -1)).toBe('0');
  });
});

test('autosum picks the contiguous numbers above the cell', () => {
  const b = autosum(fill([[0, 0, '1'], [1, 0, '2']]), 0, 2, 0)!;
  expect(raw(b, 2, 0)).toBe('=SUM(A1:A2)');
});

test('replaceAll counts replaced cells', () => {
  const res = replaceAll(fill([[0, 0, 'foo bar'], [1, 0, 'FOO']]), 0, 'foo', 'x', false);
  expect(res.count).toBe(2);
  expect(raw(res.book, 0, 0)).toBe('x bar');
});

describe('history', () => {
  test('dirty tracking across commit, undo and save', () => {
    let h = initialHistory(newBook());
    expect(isDirty(h)).toBe(false);
    h = reducer(h, { type: 'commit', fn: (b) => setCells(b, 0, [{ r: 0, c: 0, cell: { raw: '1' } }]) });
    expect(isDirty(h)).toBe(true);
    h = reducer(h, { type: 'saved', rev: h.rev });
    expect(isDirty(h)).toBe(false);
    h = reducer(h, { type: 'undo' });
    expect(isDirty(h)).toBe(true);
    expect(h.book.sheets[0].rows[0]).toBeUndefined();
  });
});
