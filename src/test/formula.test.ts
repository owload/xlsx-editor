import { describe, expect, test } from 'vitest';
import { adjustForEdit, createEvaluator, display, shiftFormula } from '../lib/formula';
import type { Book, Cell } from '../lib/types';

const book = (cells: Record<string, string>, name = 'Sheet1'): Book => {
  const rows: (Cell | undefined)[][] = [];
  for (const [addr, raw] of Object.entries(cells)) {
    const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
    const c = m[1].charCodeAt(0) - 65;
    const r = Number(m[2]) - 1;
    (rows[r] ??= [])[c] = { raw };
  }
  return { sheets: [{ name, rows, colWidths: {} }], active: 0 };
};

const calc = (cells: Record<string, string>, addr: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
  const b = book(cells);
  return createEvaluator(b).value(0, Number(m[2]) - 1, m[1].charCodeAt(0) - 65);
};

describe('evaluator', () => {
  test('arithmetic precedence and unary minus', () => {
    expect(calc({ A1: '=1+2*3' }, 'A1')).toBe(7);
    expect(calc({ A1: '=-2^2' }, 'A1')).toBe(4);
    expect(calc({ A1: '=(1+2)*3' }, 'A1')).toBe(9);
    expect(calc({ A1: '=50%' }, 'A1')).toBe(0.5);
  });

  test('references and aggregate functions', () => {
    const cells = { A1: '1', A2: '2', A3: '3', B1: '=SUM(A1:A3)', B2: '=AVERAGE(A1:A3)', B3: '=MAX(A1:A3)+MIN(A1:A3)' };
    expect(calc(cells, 'B1')).toBe(6);
    expect(calc(cells, 'B2')).toBe(2);
    expect(calc(cells, 'B3')).toBe(4);
  });

  test('accepts ";" as an argument separator', () => {
    expect(calc({ A1: '=IF(1>0;"yes";"no")' }, 'A1')).toBe('yes');
  });

  test('errors', () => {
    expect(calc({ A1: '=1/0' }, 'A1')).toEqual({ err: '#DIV/0!' });
    expect(calc({ A1: '=NOPE(1)' }, 'A1')).toEqual({ err: '#NAME?' });
    expect(calc({ A1: '=1+' }, 'A1')).toEqual({ err: '#ERROR!' });
    expect(calc({ A1: '=IFERROR(1/0;7)' }, 'A1')).toBe(7);
  });

  test('circular references do not hang', () => {
    expect(calc({ A1: '=B1', B1: '=A1' }, 'A1')).toEqual({ err: '#CIRC!' });
  });

  test('text, comparison and concatenation', () => {
    expect(calc({ A1: '="a"&"b"' }, 'A1')).toBe('ab');
    expect(calc({ A1: '="a"="A"' }, 'A1')).toBe(true);
    expect(calc({ A1: '=LEFT("hello";2)&RIGHT("hello";3)' }, 'A1')).toBe('hello');
  });

  test('cross-sheet references, including non-ASCII sheet names', () => {
    const b: Book = {
      sheets: [
        { name: 'Data', rows: [[{ raw: '5' }]], colWidths: {} },
        { name: 'Σύνολο', rows: [[{ raw: '=Data!A1*2' }], [{ raw: "=Σύνολο!A1+1" }]], colWidths: {} },
      ],
      active: 0,
    };
    const ev = createEvaluator(b);
    expect(ev.value(1, 0, 0)).toBe(10);
    expect(ev.value(1, 1, 0)).toBe(11);
  });

  test('formulas are never executed as code', () => {
    expect(calc({ A1: '=constructor.constructor("return 1")()' }, 'A1')).toEqual({ err: '#ERROR!' });
    expect(calc({ A1: '=__proto__' }, 'A1')).toEqual({ err: '#NAME?' });
  });
});

describe('formula rewriting', () => {
  test('shiftFormula moves relative references only', () => {
    expect(shiftFormula('=A1+$B$2+C$3', 2, 1)).toBe('=B3+$B$2+D$3');
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!');
    expect(shiftFormula('="A1"&A1', 1, 0)).toBe('="A1"&A2');
  });

  test('row insertion shifts references and extends ranges', () => {
    expect(adjustForEdit('=SUM(A1:A3)', 'S', 'S', 'r', 1, 1)).toBe('=SUM(A1:A4)');
    expect(adjustForEdit('=A5', 'S', 'S', 'r', 1, 2)).toBe('=A7');
  });

  test('row deletion', () => {
    expect(adjustForEdit('=A2', 'S', 'S', 'r', 1, -1)).toBe('=#REF!');
    expect(adjustForEdit('=A5', 'S', 'S', 'r', 1, -1)).toBe('=A4');
    expect(adjustForEdit('=SUM(A1:A5)', 'S', 'S', 'r', 1, -2)).toBe('=SUM(A1:A3)');
  });

  test('other sheets are left alone', () => {
    expect(adjustForEdit('=Other!A5', 'S', 'S', 'r', 0, 1)).toBe('=Other!A5');
    expect(adjustForEdit('=A5', 'S', 'Other', 'r', 0, 1)).toBe('=A5');
  });
});

describe('display', () => {
  test('numbers are trimmed of float noise', () => {
    expect(display(0.1 + 0.2)).toBe('0.3');
  });
  test('booleans, errors, empty', () => {
    expect(display(true)).toBe('TRUE');
    expect(display({ err: '#N/A' })).toBe('#N/A');
    expect(display(null)).toBe('');
  });
});
