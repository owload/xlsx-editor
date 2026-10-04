import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, test, vi } from 'vitest';
import { XlsxEditor, type XlsxEditorHandle } from '../index';
import { INDEXED_COLORS, applyTint, parseThemeColors, resolveColor, DEFAULT_THEME } from '../lib/ooxml-colors';
import { clearContents, patchStyle } from '../lib/ops';
import { layoutSheetPreview } from '../lib/preview';
import { deleteLines, insertLines } from '../lib/store';
import type { Book, Sheet } from '../lib/types';
import { inheritedStyle } from '../lib/types';
import { readWorkbook, writeWorkbook } from '../lib/xlsx-io';
import { unzip, zip } from '../lib/zip';

(globalThis as unknown as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true;

const near = (a: string, b: string, tol = 2) =>
  [1, 3, 5].every((i) => Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)) <= tol);

describe('colors stated as theme, tint or index', () => {
  test('a theme is read from its color scheme, by the index that styles use (light and dark swapped)', () => {
    const xml =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="X">' +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="112233"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="FF0000"/></a:accent1><a:accent2><a:srgbClr val="00FF00"/></a:accent2><a:accent3><a:srgbClr val="0000FF"/></a:accent3>' +
      '<a:accent4><a:srgbClr val="111111"/></a:accent4><a:accent5><a:srgbClr val="222222"/></a:accent5><a:accent6><a:srgbClr val="333333"/></a:accent6>' +
      '<a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>';
    const theme = parseThemeColors(new DOMParser().parseFromString(xml, 'application/xml'));
    expect(theme.slice(0, 5)).toEqual(['#ffffff', '#000000', '#eeeeee', '#112233', '#ff0000']);
    expect(theme[10]).toBe('#0000ee');
  });

  test('a missing or unreadable theme gives the default one', () => {
    expect(parseThemeColors(null)).toEqual(DEFAULT_THEME);
    expect(parseThemeColors(new DOMParser().parseFromString('<a/>', 'application/xml'))).toEqual(DEFAULT_THEME);
  });

  test('a tint lightens or darkens the way Excel does, in luminance', () => {
    expect(near(applyTint('#4472c4', 0.8), '#dae3f3')).toBe(true);
    expect(near(applyTint('#4472c4', 0.4), '#8faadc')).toBe(true);
    expect(near(applyTint('#4472c4', -0.25), '#2f5597')).toBe(true);
    expect(applyTint('#4472c4', 0)).toBe('#4472c4');
    expect(applyTint('#808080', 0.5)).toBe('#c0c0c0');
    expect(applyTint('#ffffff', -0.5)).toBe('#808080');
  });

  test('resolves rgb, theme with tint, and indexed colors, and refuses what cannot be shown', () => {
    expect(resolveColor({ rgb: 'FF1A2B3C' }, DEFAULT_THEME)).toBe('#1a2b3c');
    expect(resolveColor({ rgb: '1A2B3C' }, DEFAULT_THEME)).toBe('#1a2b3c');
    expect(resolveColor({ theme: '4' }, DEFAULT_THEME)).toBe('#4472c4');
    expect(near(resolveColor({ theme: '4', tint: '0.7999816888943144' }, DEFAULT_THEME)!, '#dae3f3')).toBe(true);
    expect(resolveColor({ indexed: '2' }, DEFAULT_THEME)).toBe('#ff0000');
    expect(resolveColor({ indexed: '64' }, DEFAULT_THEME)).toBeUndefined(); // automatic: no color of its own
    expect(resolveColor({ indexed: '65' }, DEFAULT_THEME)).toBeUndefined();
    expect(resolveColor({ indexed: '200' }, DEFAULT_THEME)).toBeUndefined();
    expect(resolveColor({ theme: '99' }, DEFAULT_THEME)).toBeUndefined();
    expect(resolveColor({ rgb: 'zzzzzz' }, DEFAULT_THEME)).toBeUndefined();
    expect(resolveColor({}, DEFAULT_THEME)).toBeUndefined();
    expect(INDEXED_COLORS).toHaveLength(64);
  });
});

const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const enc = (s: string) => new TextEncoder().encode(s);

/** A file made by another program: given styles and sheet XML (and optionally a theme) on top of a minimal workbook. */
async function foreign(styles: string, sheetXml: string, theme?: string): Promise<Book> {
  const files = await unzip(await writeWorkbook({ sheets: [{ name: 'S', rows: [], colWidths: {} }], active: 0 }));
  files.set('xl/styles.xml', enc(`<styleSheet ${MAIN}>${styles}</styleSheet>`));
  files.set('xl/worksheets/sheet1.xml', enc(`<worksheet ${MAIN}>${sheetXml}</worksheet>`));
  if (theme) files.set('xl/theme/theme1.xml', enc(theme));
  return readWorkbook(await zip(files));
}
const xfs = (n: number) =>
  `<cellXfs count="${n + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  Array.from({ length: n }, (_, i) => `<xf numFmtId="0" fontId="0" fillId="${i + 2}" borderId="0" applyFill="1"/>`).join('') +
  '</cellXfs>';
const fill = (inner: string) => `<fill>${inner}</fill>`;
const fills = (...f: string[]) => `<fills count="${f.length + 2}">${fill('<patternFill patternType="none"/>')}${fill('<patternFill patternType="gray125"/>')}${f.join('')}</fills>`;
const fonts = '<fonts count="1"><font><sz val="11"/></font></fonts>';
const borders = '<borders count="1"><border><left/><right/><top/><bottom/></border></borders>';

describe('fills in a file made by Excel', () => {
  test('a theme color fill, with and without a tint, an indexed fill and a gradient are shown', async () => {
    const styles =
      fonts +
      fills(
        fill('<patternFill patternType="solid"><fgColor theme="5"/></patternFill>'),
        fill('<patternFill patternType="solid"><fgColor theme="4" tint="0.7999816888943144"/></patternFill>'),
        fill('<patternFill patternType="solid"><fgColor indexed="13"/></patternFill>'),
        fill('<gradientFill degree="90"><stop position="0"><color rgb="FF00FF00"/></stop><stop position="1"><color rgb="FF0000FF"/></stop></gradientFill>'),
        fill('<patternFill patternType="solid"><fgColor rgb="FFAABBCC"/></patternFill>'),
      ) +
      borders +
      xfs(5);
    const book = await foreign(
      styles,
      '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c><c r="C1" s="3"><v>3</v></c><c r="D1" s="4"><v>4</v></c><c r="E1" s="5"><v>5</v></c></sheetData>'.replace('</sheetData>', '</row></sheetData>'),
    );
    const row = book.sheets[0].rows[0]!;
    expect(row[0]!.st!.bg).toBe('#ed7d31'); // theme 5 = accent2
    expect(near(row[1]!.st!.bg!, '#dae3f3')).toBe(true); // accent1 lightened 80%
    expect(row[2]!.st!.bg).toBe('#ffff00'); // indexed 13
    expect(row[3]!.st!.bg).toBe('#00ff00'); // the first stop of a gradient
    expect(row[4]!.st!.bg).toBe('#aabbcc');
  });

  test('the colors of the file\'s own theme are used', async () => {
    const theme =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="X"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="000000"/></a:dk2><a:lt2><a:srgbClr val="FFFFFF"/></a:lt2><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme></a:themeElements></a:theme>';
    const styles = fonts + fills(fill('<patternFill patternType="solid"><fgColor theme="4"/></patternFill>')) + borders + xfs(1);
    const book = await foreign(styles, '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>', theme);
    expect(book.sheets[0].rows[0]![0]!.st!.bg).toBe('#123456');
  });

  test('a text color given as a theme color is read too', async () => {
    const styles =
      '<fonts count="2"><font><sz val="11"/></font><font><sz val="11"/><color theme="0"/></font></fonts>' +
      fills() +
      borders +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs>';
    const book = await foreign(styles, '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData>');
    expect(book.sheets[0].rows[0]![0]!.st!.color).toBe('#ffffff');
  });
});

describe('a whole column or row that was filled', () => {
  const styles = fonts + fills(fill('<patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill>')) + borders + xfs(1);

  test('the column\'s style is read, and empty cells in it keep an explicit empty style', async () => {
    const book = await foreign(
      styles,
      '<cols><col min="2" max="2" width="12" customWidth="1" style="1"/></cols><sheetData><row r="1"><c r="B1"><v>5</v></c><c r="B2" s="0"/><c r="A1"><v>1</v></c></row></sheetData>'.replace('<c r="B2" s="0"/>', '</row><row r="2"><c r="B2" s="0"/>'),
    );
    const sheet = book.sheets[0];
    expect(sheet.colStyles).toEqual({ 1: { bg: '#ffff00' } });
    expect(sheet.rows[0]![0]!.st).toBeUndefined(); // column A has no style
    expect(sheet.rows[0]![1]!.st).toEqual({}); // B1 has a value and its own plain style: not yellow
    expect(sheet.rows[1]![1]!.st).toEqual({}); // B2 is empty, deliberately plain, and is kept
    expect(inheritedStyle(sheet, 5, 1)).toEqual({ bg: '#ffff00' }); // B6 does not exist: it is yellow
  });

  test('a cell that has the filled style itself keeps it', async () => {
    const book = await foreign(styles, '<cols><col min="2" max="2" width="9" style="1"/></cols><sheetData><row r="1"><c r="B1" s="1"><v>5</v></c></row></sheetData>');
    expect(book.sheets[0].rows[0]![1]!.st).toEqual({ bg: '#ffff00' });
  });

  test('a row style needs customFormat, and wins over the column\'s for cells that do not exist', async () => {
    const withBoth = await foreign(styles, '<cols><col min="1" max="1" width="9" style="1"/></cols><sheetData><row r="2" s="1" customFormat="1"/><row r="3" s="1"/></sheetData>');
    expect(withBoth.sheets[0].rowStyles).toEqual({ 1: { bg: '#ffff00' } });
    expect(inheritedStyle(withBoth.sheets[0], 1, 0)).toEqual({ bg: '#ffff00' });
  });

  test('is written back and read again, for columns and for rows, also when they hold no cells', async () => {
    const sheet: Sheet = { name: 'S', rows: [[{ raw: 'a' }]], colWidths: { 2: 100 }, colStyles: { 0: { bg: '#ff0000' }, 2: { bg: '#00ff00' } }, rowStyles: { 3: { bg: '#0000ff' } } };
    const bytes = await writeWorkbook({ sheets: [sheet], active: 0 });
    const xml = new TextDecoder().decode((await unzip(bytes)).get('xl/worksheets/sheet1.xml'));
    expect(xml).toMatch(/<col min="1" max="1" width="9.57" style="\d+"\/>/);
    expect(xml).toMatch(/<col min="3" max="3" width="[\d.]+" customWidth="1" style="\d+"\/>/);
    expect(xml).toMatch(/<row r="4" s="\d+" customFormat="1">/);
    const read = (await readWorkbook(bytes)).sheets[0];
    expect(read.colStyles).toEqual({ 0: { bg: '#ff0000' }, 2: { bg: '#00ff00' } });
    expect(read.rowStyles).toEqual({ 3: { bg: '#0000ff' } });
    expect(read.colWidths[2]).toBe(100);
    expect(read.rows[0]![0]!.st).toEqual({}); // A1 exists in a filled column: it is plain, and stays so
  });
});

describe('formatting a whole column, a row, or cells that do not exist', () => {
  const sheetOf = (b: Book) => b.sheets[0];
  const base = (): Book => ({ sheets: [{ name: 'S', rows: [[{ raw: 'a' }, { raw: 'b', st: { b: true } }], [undefined, { raw: 'c' }]], colWidths: {} }], active: 0 });

  test('filling whole columns sets the column style and the style of the cells that exist in them', () => {
    const b = patchStyle(base(), 0, { r1: 0, r2: 99, c1: 1, c2: 1, whole: 'col' }, { bg: '#ff0000' });
    expect(sheetOf(b).colStyles).toEqual({ 1: { bg: '#ff0000' } });
    expect(sheetOf(b).rows[0]![1]!.st).toEqual({ b: true, bg: '#ff0000' });
    expect(sheetOf(b).rows[1]![1]!.st).toEqual({ bg: '#ff0000' });
    expect(sheetOf(b).rows[0]![0]!.st).toBeUndefined(); // another column
    expect(sheetOf(b).rows.length).toBe(2); // no cell was made for the 100 rows
  });

  test('"no fill" on a whole column removes the column style and the fill of its cells', () => {
    const filled = patchStyle(base(), 0, { r1: 0, r2: 99, c1: 1, c2: 1, whole: 'col' }, { bg: '#ff0000' });
    const cleared = patchStyle(filled, 0, { r1: 0, r2: 99, c1: 1, c2: 1, whole: 'col' }, { bg: undefined });
    expect(sheetOf(cleared).colStyles).toEqual({});
    expect(sheetOf(cleared).rows[0]![1]!.st).toEqual({ b: true });
    expect(sheetOf(cleared).rows[1]![1]!.st).toBeUndefined();
  });

  test('filling a whole row works the same way', () => {
    const b = patchStyle(base(), 0, { r1: 1, r2: 1, c1: 0, c2: 25, whole: 'row' }, { bg: '#00ff00' });
    expect(sheetOf(b).rowStyles).toEqual({ 1: { bg: '#00ff00' } });
    expect(sheetOf(b).rows[1]![1]!.st).toEqual({ bg: '#00ff00' });
    expect(sheetOf(b).rows[0]![0]!.st).toBeUndefined();
  });

  test('a cell that does not exist in a filled column starts from the fill, and "no fill" there is kept as a plain cell', () => {
    const filled = patchStyle(base(), 0, { r1: 0, r2: 99, c1: 3, c2: 3, whole: 'col' }, { bg: '#ff0000' });
    const bold = patchStyle(filled, 0, { r1: 5, r2: 5, c1: 3, c2: 3 }, { b: true });
    expect(sheetOf(bold).rows[5]![3]!.st).toEqual({ bg: '#ff0000', b: true }); // changes the inherited style
    const plain = patchStyle(filled, 0, { r1: 7, r2: 7, c1: 3, c2: 3 }, { bg: undefined });
    expect(sheetOf(plain).rows[7]![3]).toEqual({ raw: '', st: {} }); // stays, so the fill does not come back
  });

  test('clearing the contents of a plain cell in a filled column does not bring the fill back', () => {
    const filled = patchStyle(base(), 0, { r1: 0, r2: 99, c1: 1, c2: 1, whole: 'col' }, { bg: '#ff0000' });
    const plain = patchStyle(filled, 0, { r1: 1, r2: 1, c1: 1, c2: 1 }, { bg: undefined });
    const cleared = clearContents(plain, 0, { r1: 1, r2: 1, c1: 1, c2: 1 });
    expect(sheetOf(cleared).rows[1]![1]).toEqual({ raw: '', st: {} });
  });
});

describe('column and row styles move with their columns and rows', () => {
  const styled = (): Book => ({
    sheets: [{ name: 'S', rows: [[{ raw: 'a' }]], colWidths: {}, colStyles: { 0: { bg: '#111111' }, 2: { bg: '#222222' } }, rowStyles: { 1: { bg: '#333333' }, 4: { bg: '#444444' } } }],
    active: 0,
  });

  test('inserting columns and rows shifts the styles', () => {
    const b = insertLines(insertLines(styled(), 0, 'c', 1, 2), 0, 'r', 2, 1);
    expect(b.sheets[0].colStyles).toEqual({ 0: { bg: '#111111' }, 4: { bg: '#222222' } });
    expect(b.sheets[0].rowStyles).toEqual({ 1: { bg: '#333333' }, 5: { bg: '#444444' } });
  });

  test('deleting columns and rows drops the styles of the deleted ones and shifts the rest', () => {
    const b = deleteLines(deleteLines(styled(), 0, 'c', 0, 1), 0, 'r', 1, 1);
    expect(b.sheets[0].colStyles).toEqual({ 1: { bg: '#222222' } });
    expect(b.sheets[0].rowStyles).toEqual({ 3: { bg: '#444444' } });
  });
});

describe('in the preview', () => {
  test('a filled column shows in the cells that do not exist, and a cell that exists covers it', () => {
    const book: Book = { sheets: [{ name: 'S', rows: [[{ raw: 'a' }, { raw: 'b' }]], colWidths: {}, colStyles: { 1: { bg: '#ffff00' } } }], active: 0 };
    const layout = layoutSheetPreview(book, 0, 360)!;
    expect(layout.fills).toHaveLength(1);
    expect(layout.fills[0]).toMatchObject({ x: layout.columns[1].x, w: layout.columns[1].w, color: '#ffff00' });
    expect(layout.cells.find((c) => c.text === 'b')!.bg).toBe('#ffffff');
    expect(layout.cells.find((c) => c.text === 'a')!.bg).toBeUndefined();
  });
});

describe('in the editor', () => {
  async function open(sheet: Partial<Sheet>) {
    const bytes = await writeWorkbook({ sheets: [{ name: 'S', rows: [[{ raw: 'x' }]], colWidths: {}, ...sheet }], active: 0 });
    const container = document.createElement('div');
    document.body.appendChild(container);
    const root = createRoot(container);
    const ref = createRef<XlsxEditorHandle>();
    const onSave = vi.fn();
    await act(async () => { root.render(createElement(XlsxEditor, { data: bytes, fileName: 'a.xlsx', onSave, ref })); });
    await act(async () => { await new Promise((r) => setTimeout(r, 80)); });
    const click = (sel: string) => act(async () => { container.querySelector<HTMLElement>(sel)!.click(); });
    const saved = async () => {
      await act(async () => { await ref.current!.save(); });
      return (await readWorkbook(onSave.mock.calls.at(-1)![0] as Uint8Array)).sheets[0];
    };
    return { container, click, saved, unmount: () => act(async () => root.unmount()) };
  }

  test('a filled column is drawn over its whole height, and the cells that exist cover it', async () => {
    const e = await open({ rows: [[{ raw: 'plain', st: {} }, { raw: 'two' }]], colStyles: { 0: { bg: '#ffff00' } } });
    const fills = [...e.container.querySelectorAll<HTMLElement>('.xe-fill')];
    expect(fills).toHaveLength(1);
    expect(['#ffff00', 'rgb(255, 255, 0)']).toContain(fills[0].style.background);
    const cell = [...e.container.querySelectorAll<HTMLElement>('.xe-cell')].find((c) => c.textContent === 'plain')!;
    expect(cell.style.background).toBe('var(--bg)'); // an explicitly plain cell is not yellow
    await e.unmount();
  });

  test('typing into an empty cell of a filled column keeps it filled', async () => {
    const e = await open({ rows: [], colStyles: { 0: { bg: '#ffff00' } } });
    const bar = e.container.querySelector<HTMLInputElement>('.xe-formulabar input')!;
    await act(async () => { bar.focus(); });
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(bar, 'typed');
      bar.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { bar.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true })); });
    const sheet = await e.saved();
    expect(sheet.rows[0]![0]).toMatchObject({ raw: 'typed', st: { bg: '#ffff00' } });
    await e.unmount();
  });

  test('choosing a column by its header and a fill color fills the whole column, not only the cells that have values', async () => {
    const e = await open({ rows: [[{ raw: 'x' }]] });
    const header = e.container.querySelectorAll<HTMLElement>('.xe-colheads .xe-hd')[1]; // column B (the test page has no layout, so few columns are drawn)
    await act(async () => { header.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); });
    await act(async () => { window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true })); });
    await e.click('button[aria-label="Fill color: more colors"]');
    await e.click('[role=menu] button[aria-label="Fill color Red"]');
    const sheet = await e.saved();
    expect(sheet.colStyles).toEqual({ 1: { bg: '#ff0000' } });
    expect(sheet.rows.length).toBe(1); // no cells were made for the rows of the column
    expect(e.container.querySelectorAll('.xe-fill').length).toBe(1);
    await e.unmount();
  });
});
