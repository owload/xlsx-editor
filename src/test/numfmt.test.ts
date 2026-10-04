import { describe, expect, test } from 'vitest';
import { formatNumber } from '../lib/numfmt';

describe('formatNumber', () => {
  test('fixed decimals and thousands separators', () => {
    expect(formatNumber('0.00', 3.14159)).toBe('3.14');
    expect(formatNumber('#,##0', 1234567)).toBe('1,234,567');
    expect(formatNumber('#,##0.00', -1234.5)).toBe('-1,234.50');
  });

  test('percent', () => {
    expect(formatNumber('0%', 0.256)).toBe('26%');
    expect(formatNumber('0.0%', 0.256)).toBe('25.6%');
  });

  test('currency with literal text', () => {
    expect(formatNumber('"$"#,##0.00', 1234.5)).toBe('$1,234.50');
  });

  test('dates and times', () => {
    expect(formatNumber('yyyy-mm-dd', 45000)).toBe('2023-03-15');
    expect(formatNumber('m/d/yyyy', 45000)).toBe('3/15/2023');
    expect(formatNumber('h:mm:ss', 0.75)).toBe('18:00:00');
    expect(formatNumber('mmm yyyy', 45000)).toBe('Mar 2023');
  });

  test('General passes through', () => {
    expect(formatNumber('General', 1.5)).toBe('1.5');
  });
});
