import { describe, expect, test } from 'vitest';
import { SHADES, STANDARD_COLORS, THEME_COLORS, collectUsedColors, normalizeHex, shade } from '../lib/colors';
import type { Book } from '../lib/types';

describe('shade', () => {
  test('mixes with white for a lighter shade and with black for a darker one', () => {
    expect(shade('#000000', 0.5)).toBe('#808080');
    expect(shade('#ffffff', -0.5)).toBe('#808080');
    expect(shade('#4472c4', 0)).toBe('#4472c4');
    expect(shade('#4472c4', 1)).toBe('#ffffff');
    expect(shade('#4472c4', -1)).toBe('#000000');
  });

  test('gives a light tint and a dark shade close to the ones Office shows for its blue', () => {
    expect(shade('#4472c4', 0.8)).toBe('#dae3f3');
    expect(shade('#4472c4', -0.25)).toBe('#335693'); // Office shows #2f5597: it mixes in HSL, this mixes in RGB
  });

  test('always returns a valid color', () => {
    for (const c of THEME_COLORS) for (const s of SHADES) expect(shade(c.hex, s.amount)).toMatch(/^#[0-9a-f]{6}$/);
  });
});

describe('the fixed palettes', () => {
  test('are ten colors wide, valid and without repeats in a row', () => {
    expect(THEME_COLORS).toHaveLength(10);
    expect(STANDARD_COLORS).toHaveLength(10);
    for (const row of [THEME_COLORS, STANDARD_COLORS]) {
      expect(new Set(row.map((c) => c.hex)).size).toBe(10);
      for (const c of row) expect(c.hex).toMatch(/^#[0-9a-f]{6}$/);
    }
  });

  test('have five shades under the theme colors', () => {
    expect(SHADES.map((s) => s.amount)).toEqual([0.8, 0.6, 0.4, -0.25, -0.5]);
  });
});

describe('normalizeHex', () => {
  test.each([
    ['ff8800', '#ff8800'],
    ['#FF8800', '#ff8800'],
    ['  #ff8800  ', '#ff8800'],
    ['#f80', '#ff8800'],
    ['ABC', '#aabbcc'],
  ])('%s -> %s', (input, out) => {
    expect(normalizeHex(input)).toBe(out);
  });

  test.each(['', '#', 'ff88', 'ff88000', 'gggggg', '#12345', 'red', 'rgb(1,2,3)'])('%s is not a color', (input) => {
    expect(normalizeHex(input)).toBeNull();
  });
});

describe('collectUsedColors', () => {
  const cell = (st: object) => ({ raw: 'x', st });
  const book = (...rows: (object | undefined)[][]): Book => ({
    sheets: [{ name: 'S', rows: rows.map((r) => r.map((st) => (st ? cell(st) : undefined))), colWidths: {} }],
    active: 0,
  });

  test('lists fill and text colors, the most used first, once each, in lower case', () => {
    const b = book([{ bg: '#FF0000' }, { bg: '#00ff00' }, { bg: '#ff0000', color: '#0000ff' }], [{ color: '#00FF00' }, { bg: '#ff0000' }]);
    expect(collectUsedColors(b)).toEqual(['#ff0000', '#00ff00', '#0000ff']);
  });

  test('keeps the order of appearance for colors used equally often', () => {
    const b = book([{ bg: '#111111' }, { bg: '#222222' }, { bg: '#333333' }]);
    expect(collectUsedColors(b)).toEqual(['#111111', '#222222', '#333333']);
  });

  test('stops at the limit and reads every sheet', () => {
    const many = book(Array.from({ length: 15 }, (_, i) => ({ bg: '#' + (i + 1).toString(16).padStart(2, '0').repeat(3) })));
    expect(collectUsedColors(many)).toHaveLength(10);
    expect(collectUsedColors(many, 3)).toHaveLength(3);
    const two: Book = { sheets: [book([{ bg: '#aaaaaa' }]).sheets[0], book([{ color: '#bbbbbb' }]).sheets[0]], active: 0 };
    expect(collectUsedColors(two)).toEqual(['#aaaaaa', '#bbbbbb']);
  });

  test('ignores invalid colors and cells without style, and an empty workbook has none', () => {
    expect(collectUsedColors(book([{ bg: 'red' }, undefined, { color: 'not a color' }, {}]))).toEqual([]);
    expect(collectUsedColors({ sheets: [{ name: 'S', rows: [], colWidths: {} }], active: 0 })).toEqual([]);
  });
});
