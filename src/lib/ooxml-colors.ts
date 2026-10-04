/**
 * Colors the way an .xlsx file states them: an RGB value, a theme color with a tint, or an index into the
 * old fixed palette. They are turned into plain "#rrggbb" values when a file is read, so a fill or a text
 * color that Excel shows in a theme color looks the same here.
 */

/** The theme colors by the index that cell styles use: 0 light 1, 1 dark 1, 2 light 2, 3 dark 2, 4-9 accents, 10-11 links. */
export type ThemeColors = string[];

/** The default Office theme, used when a file has none or it cannot be read. */
export const DEFAULT_THEME: ThemeColors = [
  '#ffffff', // lt1
  '#000000', // dk1
  '#e7e6e6', // lt2
  '#44546a', // dk2
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
  '#0563c1', // hyperlink
  '#954f72', // followed hyperlink
];

/** The order the colors are written in a theme's color scheme. */
const SCHEME_ORDER = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
/** Where each of them goes in the index used by styles (the first two and the next two are swapped). */
const THEME_INDEX = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];

const HEX6 = /^[0-9a-fA-F]{6}$/;

/** The colors of the first `a:clrScheme` of a theme part, by style index; the default theme if it cannot be read. */
export function parseThemeColors(doc: Document | null): ThemeColors {
  const out = DEFAULT_THEME.slice();
  const scheme = doc ? Array.from(doc.getElementsByTagName('*')).find((e) => e.localName === 'clrScheme') : undefined;
  if (!scheme) return out;
  for (const el of Array.from(scheme.children)) {
    const slot = SCHEME_ORDER.indexOf(el.localName);
    if (slot < 0) continue;
    const child = el.firstElementChild;
    const value = child?.getAttribute('val') ?? child?.getAttribute('lastClr');
    if (value && HEX6.test(value)) out[THEME_INDEX[slot]] = '#' + value.toLowerCase();
  }
  return out;
}

/** The fixed palette that `indexed` colors refer to (the 64 colors of the old spreadsheet palette). */
export const INDEXED_COLORS: string[] = (
  '000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff ' +
  '000000 ffffff ff0000 00ff00 0000ff ffff00 ff00ff 00ffff ' +
  '800000 008000 000080 808000 800080 008080 c0c0c0 808080 ' +
  '9999ff 993366 ffffcc ccffff 660066 ff8080 0066cc ccccff ' +
  '000080 ff00ff ffff00 00ffff 800080 800000 008080 0000ff ' +
  '00ccff ccffff ccffcc ffff99 99ccff ff99cc cc99ff ffcc99 ' +
  '3366ff 33ccff 99cc00 ffcc00 ff9900 ff6600 666699 969696 ' +
  '003366 339966 003300 333300 993300 993366 333399 333333'
)
  .split(' ')
  .map((h) => '#' + h);

const toHex = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');

/** Excel's tint: lightens (positive) or darkens (negative) a color by changing its luminance in HSL. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  l = Math.max(0, Math.min(1, l));
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) return '#' + toHex(l * 255).repeat(3);
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return '#' + [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)].map((v) => toHex(v * 255)).join('');
}

export interface ColorAttrs {
  rgb?: string | null;
  theme?: string | null;
  tint?: string | null;
  indexed?: string | null;
}

/**
 * "#rrggbb" for the color an element states, or undefined if it has none that can be shown (automatic, an
 * unknown index, anything malformed). `rgb` is "AARRGGBB" or "RRGGBB"; the alpha is ignored.
 */
export function resolveColor(a: ColorAttrs, theme: ThemeColors): string | undefined {
  if (a.rgb) {
    const v = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb;
    return HEX6.test(v) ? '#' + v.toLowerCase() : undefined;
  }
  if (a.theme != null && a.theme !== '') {
    const base = theme[Number(a.theme)];
    if (!base) return undefined;
    const tint = Number(a.tint ?? 0);
    return applyTint(base, Number.isFinite(tint) ? Math.max(-1, Math.min(1, tint)) : 0);
  }
  if (a.indexed != null && a.indexed !== '') {
    const i = Number(a.indexed);
    // 64 and 65 are the system foreground and background colors: black and white.
    if (i === 64) return '#000000';
    if (i === 65) return '#ffffff';
    return INDEXED_COLORS[i];
  }
  return undefined;
}
