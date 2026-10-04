import { unzip } from './zip';

/**
 * What saving a workbook with this editor would drop. The editor rebuilds the file from the data it
 * understands (values, formulas, number formats, basic cell formatting, column widths, merged cells,
 * sheets), so everything else in the file is lost; this lists the notable things that are there.
 * It only reads the file and applies the same size limits as opening it.
 *
 * Plain fonts and other cosmetic details are not listed: nearly every file has them, and a note
 * that always appears would be ignored.
 */

export interface LossFinding {
  /** Stable identifier, for tests ("charts"). */
  id: string;
  /** Short English text shown to the user ("Charts"). */
  label: string;
}

export interface InspectResult {
  /** Empty when nothing would be lost. */
  unsupported: LossFinding[];
}

const dec = new TextDecoder();

interface Rule {
  id: string;
  label: string;
  /** True if the file has the feature. */
  test(names: string[], texts: { workbook: string; styles: string; strings: string; sheets: string[] }): boolean;
}

const hasPart = (names: string[], prefix: string) => names.some((n) => n.startsWith(prefix));
const inSheets = (sheets: string[], re: RegExp) => sheets.some((s) => re.test(s));

const RULES: Rule[] = [
  { id: 'charts', label: 'Charts', test: (n) => hasPart(n, 'xl/charts/') },
  { id: 'images', label: 'Images', test: (n) => hasPart(n, 'xl/media/') },
  {
    id: 'drawings',
    label: 'Shapes and drawings',
    test: (n) => hasPart(n, 'xl/drawings/') && !hasPart(n, 'xl/charts/') && !hasPart(n, 'xl/media/'),
  },
  { id: 'comments', label: 'Comments', test: (n) => n.some((x) => /^xl\/(comments|threadedComments)/.test(x)) },
  { id: 'pivotTables', label: 'Pivot tables', test: (n) => hasPart(n, 'xl/pivotTables/') },
  { id: 'tables', label: 'Tables', test: (n) => hasPart(n, 'xl/tables/') },
  { id: 'macros', label: 'Macros', test: (n) => n.includes('xl/vbaProject.bin') },
  { id: 'externalLinks', label: 'Links to other workbooks', test: (n) => hasPart(n, 'xl/externalLinks/') },
  { id: 'conditionalFormatting', label: 'Conditional formatting', test: (_n, t) => inSheets(t.sheets, /<conditionalFormatting[\s>]/) },
  { id: 'dataValidation', label: 'Data validation', test: (_n, t) => inSheets(t.sheets, /<dataValidations[\s>]/) },
  { id: 'freezePanes', label: 'Frozen panes', test: (_n, t) => inSheets(t.sheets, /<pane\b[^>]*\b(xSplit|ySplit)=/) },
  { id: 'rowHeights', label: 'Row heights', test: (_n, t) => inSheets(t.sheets, /<row\b[^>]*\bcustomHeight=["']?(1|true)/) },
  { id: 'hyperlinks', label: 'Hyperlinks', test: (_n, t) => inSheets(t.sheets, /<hyperlinks[\s>]/) },
  { id: 'filters', label: 'Filters', test: (_n, t) => inSheets(t.sheets, /<autoFilter[\s>]/) },
  { id: 'protection', label: 'Sheet protection', test: (_n, t) => inSheets(t.sheets, /<sheetProtection[\s>]/) },
  {
    id: 'arrayFormulas',
    label: 'Array formulas',
    test: (_n, t) => inSheets(t.sheets, /<f\b[^>]*\bt=["']array["']/),
  },
  {
    id: 'sharedFormulas',
    label: 'Shared formulas (kept as values)',
    test: (_n, t) => inSheets(t.sheets, /<f\b[^>]*\bt=["']shared["']/),
  },
  {
    id: 'hidden',
    label: 'Hidden sheets, rows or columns',
    test: (_n, t) =>
      /<sheet\b[^>]*\bstate=["'](hidden|veryHidden)["']/.test(t.workbook) ||
      inSheets(t.sheets, /<(row|col)\b[^>]*\bhidden=["']?(1|true)/),
  },
  { id: 'definedNames', label: 'Named ranges', test: (_n, t) => /<definedName[\s>]/.test(t.workbook) },
  { id: 'themeColors', label: 'Theme colors', test: (_n, t) => /<(color|fgColor|bgColor)\b[^>]*\btheme=/.test(t.styles) },
  { id: 'richText', label: 'Formatting inside cells', test: (_n, t) => /<r[\s>]/.test(t.strings) },
];

export async function inspectWorkbook(data: Uint8Array): Promise<InspectResult> {
  const files = await unzip(data);
  const names = [...files.keys()];
  const text = (name: string) => (files.has(name) ? dec.decode(files.get(name)) : '');
  const texts = {
    workbook: text('xl/workbook.xml'),
    styles: text('xl/styles.xml'),
    strings: text('xl/sharedStrings.xml'),
    sheets: names.filter((n) => /^xl\/worksheets\/[^/]+\.xml$/.test(n)).map(text),
  };
  const unsupported: LossFinding[] = RULES.filter((rule) => rule.test(names, texts)).map(({ id, label }) => ({ id, label }));
  return { unsupported };
}
