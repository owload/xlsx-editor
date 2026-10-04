import { describe, expect, test } from 'vitest';
import { inspectWorkbook } from '../lib/inspect';
import { writeWorkbook } from '../lib/xlsx-io';
import { zip } from '../lib/zip';

const enc = new TextEncoder();
const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

/** A workbook file made of the given parts, on top of the minimal ones the reader needs. */
async function file(parts: Record<string, string>): Promise<Uint8Array> {
  const base: Record<string, string> = {
    'xl/workbook.xml': `<workbook ${MAIN}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/worksheets/sheet1.xml': `<worksheet ${MAIN}><sheetData/></worksheet>`,
  };
  return zip(new Map(Object.entries({ ...base, ...parts }).map(([k, v]) => [k, enc.encode(v)])));
}
const ids = async (data: Uint8Array) => (await inspectWorkbook(data)).unsupported.map((u) => u.id);
const sheet = (inner: string, attrs = '') => ({ 'xl/worksheets/sheet1.xml': `<worksheet ${MAIN}>${inner}</worksheet>${attrs}` });

describe('inspectWorkbook', () => {
  test('a file the editor wrote itself loses nothing', async () => {
    const own = await writeWorkbook({
      sheets: [{ name: 'Data', rows: [[{ raw: 'a', st: { b: true, bg: '#ffff00' } }, { raw: '=1+2' }]], colWidths: { 0: 120 }, merges: [{ s: { r: 0, c: 3 }, e: { r: 0, c: 4 } }] }],
      active: 0,
    });
    expect(await ids(own)).toEqual([]);
  });

  test('a plain file has nothing to report', async () => {
    expect(await ids(await file({}))).toEqual([]);
  });

  test.each([
    ['charts', { 'xl/charts/chart1.xml': '<c/>' }],
    ['images', { 'xl/media/image1.png': 'x' }],
    ['drawings', { 'xl/drawings/drawing1.xml': '<d/>' }],
    ['comments', { 'xl/comments1.xml': '<c/>' }],
    ['comments', { 'xl/threadedComments/threadedComment1.xml': '<c/>' }],
    ['pivotTables', { 'xl/pivotTables/pivotTable1.xml': '<p/>' }],
    ['tables', { 'xl/tables/table1.xml': '<t/>' }],
    ['macros', { 'xl/vbaProject.bin': 'x' }],
    ['externalLinks', { 'xl/externalLinks/externalLink1.xml': '<e/>' }],
  ])('finds %s from the parts of the archive', async (id, parts) => {
    expect(await ids(await file(parts))).toEqual([id]);
  });

  test('a chart or an image does not also report "drawings", which only hold them', async () => {
    const found = await ids(await file({ 'xl/drawings/drawing1.xml': '<d/>', 'xl/charts/chart1.xml': '<c/>', 'xl/media/i.png': 'x' }));
    expect(found).toEqual(['charts', 'images']);
  });

  test.each([
    ['conditionalFormatting', '<conditionalFormatting sqref="A1"><cfRule/></conditionalFormatting><sheetData/>'],
    ['dataValidation', '<sheetData/><dataValidations count="1"><dataValidation/></dataValidations>'],
    ['freezePanes', '<sheetViews><sheetView><pane xSplit="1" ySplit="1" state="frozen"/></sheetView></sheetViews><sheetData/>'],
    ['hyperlinks', '<sheetData/><hyperlinks><hyperlink ref="A1"/></hyperlinks>'],
    ['filters', '<sheetData/><autoFilter ref="A1:B2"/>'],
    ['protection', '<sheetData/><sheetProtection sheet="1"/>'],
    ['arrayFormulas', '<sheetData><row r="1"><c r="A1"><f t="array" ref="A1:A2">SUM(B1)</f></c></row></sheetData>'],
    ['sharedFormulas', '<sheetData><row r="1"><c r="A1"><f t="shared" si="0">B1</f></c></row></sheetData>'],
    ['hidden', '<sheetData><row r="2" hidden="1"/></sheetData>'],
    ['hidden', '<cols><col min="1" max="1" hidden="1"/></cols><sheetData/>'],
  ])('finds %s in a sheet', async (id, inner) => {
    expect(await ids(await file(sheet(inner)))).toEqual([id]);
  });

  test('ordinary row attributes are not reported, and neither are row heights, which the editor keeps', async () => {
    expect(await ids(await file(sheet('<sheetData><row r="1" ht="15" spans="1:2"/></sheetData>')))).toEqual([]);
    expect(await ids(await file(sheet('<sheetData><row r="1" ht="30" customHeight="1"/></sheetData>')))).toEqual([]);
  });

  test('finds things in the workbook, styles and shared strings', async () => {
    expect(await ids(await file({ 'xl/workbook.xml': `<workbook ${MAIN}><sheets><sheet name="S" sheetId="1" state="hidden"/></sheets></workbook>` }))).toEqual(['hidden']);
    expect(await ids(await file({ 'xl/workbook.xml': `<workbook ${MAIN}><sheets/><definedNames><definedName name="X">S!A1</definedName></definedNames></workbook>` }))).toEqual(['definedNames']);
    // Theme colors are turned into fixed colors on opening, so they look the same and are not reported.
    expect(await ids(await file({ 'xl/styles.xml': `<styleSheet ${MAIN}><fonts><font><color theme="1"/></font></fonts></styleSheet>` }))).toEqual([]);
    expect(await ids(await file({ 'xl/styles.xml': `<styleSheet ${MAIN}><fonts><font><color rgb="FF000000"/></font></fonts></styleSheet>` }))).toEqual([]);
    expect(await ids(await file({ 'xl/sharedStrings.xml': `<sst ${MAIN}><si><r><rPr/><t>a</t></r><r><t>b</t></r></si></sst>` }))).toEqual(['richText']);
    expect(await ids(await file({ 'xl/sharedStrings.xml': `<sst ${MAIN}><si><t>plain</t></si></sst>` }))).toEqual([]);
  });

  test('reports several features, each once, with a stable id and a label', async () => {
    const result = await inspectWorkbook(await file({
      'xl/charts/chart1.xml': '<c/>',
      'xl/comments1.xml': '<c/>',
      ...sheet('<conditionalFormatting sqref="A1"/><conditionalFormatting sqref="B1"/><sheetData/>'),
    }));
    expect(result.unsupported).toEqual([
      { id: 'charts', label: 'Charts' },
      { id: 'comments', label: 'Comments' },
      { id: 'conditionalFormatting', label: 'Conditional formatting' },
    ]);
  });

  test('refuses something that is not a workbook archive', async () => {
    await expect(inspectWorkbook(enc.encode('not a zip'))).rejects.toThrow();
  });
});
