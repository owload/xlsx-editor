/** Built-in Excel number formats by id (ECMA-376, 18.8.30). */
export const BUILTIN_FMT: Record<number, string> = {
  1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0', 38: '#,##0', 39: '#,##0.00', 40: '#,##0.00', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0',
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

const clean = (s: string) => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_./g, ' ').replace(/\*./g, '');
const stripMeta = (s: string) => s.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '').replace(/\\./g, '');

export function isDateFormat(z: string): boolean {
  return /[ymdhs]/i.test(stripMeta(z).replace(/General/gi, ''));
}

const pad = (n: number, w = 2) => String(n).padStart(w, '0');

function formatDate(z: string, v: number): string | null {
  if (v < 0 || v > 2958465) return null;
  const d = new Date(Math.round((v - 25569) * 86400000));
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate();
  const h = d.getUTCHours(), mi = d.getUTCMinutes(), s = d.getUTCSeconds();
  const ampm = /AM\/PM|A\/P/i.test(z);
  const parts = z.split(/("[^"]*"|\\.|\[[^\]]*\])/).map((seg, i) => {
    if (i % 2) return seg.startsWith('"') ? seg.slice(1, -1) : seg.startsWith('\\') ? seg[1] : '';
    return seg;
  });
  let out = '';
  for (let i = 0; i < parts.length; i++) {
    if (i % 2) { out += parts[i]; continue; }
    const seg = parts[i];
    const re = /yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|AM\/PM|A\/P|\.0+/gi;
    let last = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(seg))) {
      out += seg.slice(last, m.index);
      last = m.index + m[0].length;
      const t = m[0].toLowerCase();
      const before = (out + seg.slice(0, m.index)).toLowerCase();
      const afterTxt = seg.slice(last).toLowerCase();
      const isMin = t[0] === 'm' && t.length <= 2 && (/h[^a-z]*$/.test(before) || /^[^a-z]*s/.test(afterTxt));
      if (t === 'yyyy') out += pad(Y, 4);
      else if (t === 'yy') out += pad(Y % 100);
      else if (t === 'mmmm') out += MONTHS[M];
      else if (t === 'mmm') out += MONTHS[M].slice(0, 3);
      else if (t === 'mm') out += isMin ? pad(mi) : pad(M + 1);
      else if (t === 'm') out += isMin ? mi : M + 1;
      else if (t === 'dd') out += pad(D);
      else if (t === 'd') out += D;
      else if (t === 'hh') out += pad(ampm ? h % 12 || 12 : h);
      else if (t === 'h') out += ampm ? h % 12 || 12 : h;
      else if (t === 'ss') out += pad(s);
      else if (t === 's') out += s;
      else if (t[0] === '.') out += '.' + pad(d.getUTCMilliseconds(), 3).slice(0, t.length - 1);
      else out += h < 12 ? 'AM' : 'PM';
    }
    out += seg.slice(last);
  }
  return out;
}

function formatPlain(z: string, v: number): string | null {
  const m = /[#0?]+[#0?,]*(\.[0#?]*)?|\.[0#?]+/.exec(z);
  if (!m) return null;
  const pat = m[0];
  const prefix = clean(z.slice(0, m.index));
  const suffix = clean(z.slice(m.index + pat.length));
  const percents = (z.match(/%/g) ?? []).length;
  const dot = pat.indexOf('.');
  const intPat = dot < 0 ? pat : pat.slice(0, dot);
  const decimals = dot < 0 ? 0 : pat.length - dot - 1;
  const grouping = intPat.includes(',');
  const minInt = (intPat.match(/0/g) ?? []).length;
  const x = Math.abs(v) * Math.pow(100, percents);
  if (/e[+-]/i.test(z)) return v.toExponential(decimals).replace('e', 'E');
  const [i0, f = ''] = x.toFixed(Math.min(decimals, 20)).split('.');
  let i = i0;
  if (i.length < minInt) i = i.padStart(minInt, '0');
  if (minInt === 0 && i === '0' && decimals > 0) i = '';
  if (grouping) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const body = i + (decimals > 0 ? '.' + f : '');
  return (v < 0 ? '-' : '') + prefix + body + suffix;
}

/** Formats a number with an Excel format code. Returns null if the format is unsupported. */
export function formatNumber(z: string, v: number): string | null {
  if (z === '@') return String(v);
  const sections = z.split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  let sec = sections[0];
  if (v < 0 && sections.length > 1) {
    sec = sections[1];
    v = -v;
  } else if (v === 0 && sections.length > 2) sec = sections[2];
  if (/^\s*(\[[^\]]*\])*General(\s*\[[^\]]*\])*\s*$/i.test(sec)) return String(+v.toPrecision(12));
  try {
    return isDateFormat(sec) ? formatDate(sec, v) : formatPlain(sec, v);
  } catch {
    return null;
  }
}
