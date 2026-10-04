import type { Book } from './types';

/** The colors of the fill and text color menus, in the way spreadsheet programs offer them. */

export interface NamedColor {
  hex: string;
  name: string;
}

/** The first row of the theme colors (the base colors of the Office palette). */
export const THEME_COLORS: NamedColor[] = [
  { hex: '#ffffff', name: 'White' },
  { hex: '#000000', name: 'Black' },
  { hex: '#e7e6e6', name: 'Light Gray' },
  { hex: '#44546a', name: 'Blue-Gray' },
  { hex: '#4472c4', name: 'Blue' },
  { hex: '#ed7d31', name: 'Orange' },
  { hex: '#a5a5a5', name: 'Gray' },
  { hex: '#ffc000', name: 'Gold' },
  { hex: '#5b9bd5', name: 'Light Blue' },
  { hex: '#70ad47', name: 'Green' },
];

export const STANDARD_COLORS: NamedColor[] = [
  { hex: '#c00000', name: 'Dark Red' },
  { hex: '#ff0000', name: 'Red' },
  { hex: '#ffc000', name: 'Orange' },
  { hex: '#ffff00', name: 'Yellow' },
  { hex: '#92d050', name: 'Light Green' },
  { hex: '#00b050', name: 'Green' },
  { hex: '#00b0f0', name: 'Light Blue' },
  { hex: '#0070c0', name: 'Blue' },
  { hex: '#002060', name: 'Dark Blue' },
  { hex: '#7030a0', name: 'Purple' },
];

/** Lighter (positive) and darker (negative) shades shown under each theme color, as a fraction. */
export const SHADES: { amount: number; label: string }[] = [
  { amount: 0.8, label: 'Lighter 80%' },
  { amount: 0.6, label: 'Lighter 60%' },
  { amount: 0.4, label: 'Lighter 40%' },
  { amount: -0.25, label: 'Darker 25%' },
  { amount: -0.5, label: 'Darker 50%' },
];

const channel = (hex: string, i: number) => parseInt(hex.slice(1 + i * 2, 3 + i * 2), 16);
const toHex = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');

/** A shade of `hex`: mixed with white by `amount` (0..1), or with black by `-amount` (-1..0). */
export function shade(hex: string, amount: number): string {
  const mix = (c: number) => (amount >= 0 ? c + (255 - c) * amount : c * (1 + amount));
  return '#' + [0, 1, 2].map((i) => toHex(mix(channel(hex, i)))).join('');
}

/** "#RRGGBB" in lower case from what a person types ("ff8800", "#F80", " #FF8800 "); null if it is not a color. */
export function normalizeHex(input: string): string | null {
  const s = input.trim().replace(/^#/, '').toLowerCase();
  if (/^[0-9a-f]{3}$/.test(s)) return '#' + [...s].map((c) => c + c).join('');
  if (/^[0-9a-f]{6}$/.test(s)) return '#' + s;
  return null;
}

/**
 * The colors the workbook already uses for fills and text, the most used first (ties keep the order of
 * appearance), without repeats and at most `limit` of them.
 */
export function collectUsedColors(book: Book, limit = 10): string[] {
  const counts = new Map<string, number>();
  for (const sheet of book.sheets) {
    for (const row of sheet.rows) {
      if (!row) continue;
      for (const cell of row) {
        for (const c of [cell?.st?.bg, cell?.st?.color]) {
          const hex = c && normalizeHex(c);
          if (hex) counts.set(hex, (counts.get(hex) ?? 0) + 1);
        }
      }
    }
  }
  return [...counts.entries()]
    .map(([hex, n], order) => ({ hex, n, order }))
    .sort((a, b) => b.n - a.n || a.order - b.order)
    .slice(0, limit)
    .map((e) => e.hex);
}
