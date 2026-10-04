import { describe, expect, test } from 'vitest';
import { readWorkbook, writeWorkbook } from '../lib/xlsx-io';
import { createEvaluator, display } from '../lib/formula';
import { zip } from '../lib/zip';
import type { Book } from '../lib/types';

// The reader uses DOMParser, which the test environment (happy-dom) provides.
describe('xlsx read/write', () => {
  const sample = (): Book => ({
    active: 0,
    sheets: [
      {
        name: 'Data & <Co>',
        colWidths: { 1: 120 },
        rows: [
          [{ raw: 'Name', st: { b: true, bg: '#217346', color: '#ffffff', h: 'center', bd: true } }, { raw: 'Qty' }],
          [{ raw: 'Apples' }, { raw: '3' }],
          [{ raw: '007', text: true }, { raw: '=B2*2', z: '0.00', st: { i: true } }],
        ],
      },
      { name: 'Second', colWidths: {}, rows: [[{ raw: "=SUM('Data & <Co>'!B2:B3)" }]] },
    ],
  });

  test('round-trips values, formulas, formats, styles and widths', async () => {
    const bytes = await writeWorkbook(sample());
    const back = await readWorkbook(bytes);
    expect(back.sheets.map((s) => s.name)).toEqual(['Data & <Co>', 'Second']);
    expect(back.sheets[0].rows[0]![0]!.st).toEqual({ b: true, color: '#ffffff', bg: '#217346', bd: true, h: 'center' });
    expect(back.sheets[0].rows[2]![0]).toMatchObject({ raw: '007', text: true });
    expect(back.sheets[0].rows[2]![1]).toMatchObject({ raw: '=B2*2', z: '0.00' });
    expect(back.sheets[0].colWidths[1]).toBe(120);
    const ev = createEvaluator(back);
    expect(display(ev.value(0, 2, 1), '0.00')).toBe('6.00');
    expect(ev.value(1, 0, 0)).toBe(9);
  });

  test('rejects XML with a DOCTYPE', async () => {
    const enc = new TextEncoder();
    const evil = await zip(new Map([['xl/workbook.xml', enc.encode('<!DOCTYPE x [<!ENTITY a "b">]><workbook/>')]]));
    await expect(readWorkbook(evil)).rejects.toThrow(/DTD|ENTITY/);
  });
});
