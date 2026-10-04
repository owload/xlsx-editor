import type { Book, Cell, Range, Sheet, Style } from './types';
import { isColor } from './types';
import { usedSize } from './types';
import { decodeAddr, encodeAddr, encodeCol, decodeCol } from './addr';
import { createEvaluator, isErr, parseLiteral } from './formula';
import { BUILTIN_FMT } from './numfmt';
import { parseThemeColors, resolveColor } from './ooxml-colors';
import { DEFAULT_ROW_HEIGHT, MAX_ROW_HEIGHT, MIN_ROW_HEIGHT } from './store';
import { unzip, zip } from './zip';

const ERR_CODES = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];

function xml(bytes: Uint8Array | undefined, name: string): Document {
  if (!bytes) throw new Error(`Missing ${name} in the file`);
  const text = new TextDecoder().decode(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('DTD/ENTITY in XML is not allowed');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error(`Invalid XML: ${name}`);
  return doc;
}

const kids = (el: Element | null | undefined, tag: string): Element[] =>
  el ? Array.from(el.children).filter((c) => c.localName === tag) : [];
const kid = (el: Element | null | undefined, tag: string) => kids(el, tag)[0];

function richText(si: Element): string {
  let s = '';
  for (const t of Array.from(si.getElementsByTagName('t'))) {
    // skip phonetic hints <rPh>
    if (t.parentElement?.localName === 'rPh') continue;
    s += t.textContent ?? '';
  }
  return s;
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export async function readWorkbook(buf: ArrayBuffer | Uint8Array): Promise<Book> {
  const files = await unzip(buf);
  const wb = xml(files.get('xl/workbook.xml'), 'xl/workbook.xml').documentElement;
  const rels = xml(files.get('xl/_rels/workbook.xml.rels'), 'workbook.xml.rels').documentElement;
  const relTarget = new Map<string, string>();
  for (const r of Array.from(rels.children)) {
    const id = r.getAttribute('Id');
    const t = r.getAttribute('Target');
    if (id && t) relTarget.set(id, resolvePath('xl/workbook.xml', t));
  }

  const shared: string[] = [];
  const ssBytes = files.get('xl/sharedStrings.xml');
  if (ssBytes) for (const si of kids(xml(ssBytes, 'sharedStrings').documentElement, 'si')) shared.push(richText(si));

  // style index (xf) -> number format and formatting
  const numFmts = new Map<number, string>();
  const xfs: { z?: string; st?: Style }[] = [];
  const stBytes = files.get('xl/styles.xml');
  if (stBytes) {
    const st = xml(stBytes, 'styles').documentElement;
    for (const nf of kids(kid(st, 'numFmts'), 'numFmt')) {
      numFmts.set(Number(nf.getAttribute('numFmtId')), nf.getAttribute('formatCode') ?? '');
    }
    const flag = (el: Element | undefined) => !!el && !['0', 'false'].includes(el.getAttribute('val') ?? '1');
    // Colors may be stated as RGB, as a theme color with a tint, or by an index into the old palette.
    const themePart = [...files.keys()].find((k) => /^xl\/theme\/[^/]+\.xml$/.test(k));
    let themeDoc: Document | null = null;
    try {
      themeDoc = themePart ? xml(files.get(themePart), themePart) : null;
    } catch {
      themeDoc = null; // an unreadable theme only means the default colors are used
    }
    const theme = parseThemeColors(themeDoc);
    const rgb = (el: Element | undefined | null) =>
      el
        ? resolveColor(
            { rgb: el.getAttribute('rgb'), theme: el.getAttribute('theme'), tint: el.getAttribute('tint'), indexed: el.getAttribute('indexed') },
            theme,
          )
        : undefined;
    const fonts = kids(kid(st, 'fonts'), 'font').map((f) => {
      const o: Style = {};
      if (flag(kid(f, 'b'))) o.b = true;
      if (flag(kid(f, 'i'))) o.i = true;
      if (kid(f, 'u') && kid(f, 'u')!.getAttribute('val') !== 'none') o.u = true;
      const sz = Number(kid(f, 'sz')?.getAttribute('val'));
      if (sz >= 6 && sz <= 72 && sz !== 11) o.sz = sz;
      const col = rgb(kid(f, 'color'));
      if (col && col !== '#000000') o.color = col;
      return o;
    });
    const fills = kids(kid(st, 'fills'), 'fill').map((f) => {
      const pf = kid(f, 'patternFill');
      if (pf) return pf.getAttribute('patternType') === 'solid' ? rgb(kid(pf, 'fgColor')) : undefined;
      // A gradient is shown in the color it starts with.
      const first = kid(kid(f, 'gradientFill'), 'stop');
      return rgb(kid(first, 'color'));
    });
    const borders = kids(kid(st, 'borders'), 'border').map((b) => !!kid(b, 'left')?.getAttribute('style') && !!kid(b, 'bottom')?.getAttribute('style'));
    for (const xf of kids(kid(st, 'cellXfs'), 'xf')) {
      const id = Number(xf.getAttribute('numFmtId') ?? 0);
      const f = numFmts.get(id) ?? BUILTIN_FMT[id];
      const style: Style = { ...fonts[Number(xf.getAttribute('fontId') ?? 0)] };
      const bg = fills[Number(xf.getAttribute('fillId') ?? 0)];
      if (bg) style.bg = bg;
      if (borders[Number(xf.getAttribute('borderId') ?? 0)]) style.bd = true;
      const h = kid(xf, 'alignment')?.getAttribute('horizontal');
      if (h === 'left' || h === 'center' || h === 'right') style.h = h;
      xfs.push({ z: f && f !== 'General' ? f : undefined, st: Object.keys(style).length ? style : undefined });
    }
  }

  const sheets: Sheet[] = [];
  for (const sh of kids(kid(wb, 'sheets'), 'sheet')) {
    const rid = Array.from(sh.attributes).find((a) => a.localName === 'id' || a.name.endsWith(':id'))?.value;
    const path = rid ? relTarget.get(rid) : undefined;
    const ws = xml(path ? files.get(path) : undefined, path ?? 'sheet').documentElement;
    const sheet: Sheet = { name: sh.getAttribute('name') ?? `Sheet${sheets.length + 1}`, rows: [], colWidths: {} };

    for (const col of kids(kid(ws, 'cols'), 'col')) {
      const w = Number(col.getAttribute('width'));
      const min = Number(col.getAttribute('min'));
      const max = Math.min(Number(col.getAttribute('max')), min + 500);
      if (w > 0 && min > 0) for (let c = min; c <= max; c++) sheet.colWidths[c - 1] = Math.round(w * 7 + 5);
      // A column that was filled or formatted as a whole: the cells that do not exist have this style.
      const colStyle = xfs[Number(col.getAttribute('style') ?? 0)]?.st;
      if (colStyle && min > 0) for (let c = min; c <= max; c++) (sheet.colStyles ??= {})[c - 1] = colStyle;
    }
    const merges: Range[] = [];
    for (const mc of kids(kid(ws, 'mergeCells'), 'mergeCell')) {
      const [a, b] = (mc.getAttribute('ref') ?? '').split(':');
      const s = decodeAddr(a ?? '');
      const e = decodeAddr(b ?? '');
      if (s && e) merges.push({ s, e });
    }
    if (merges.length) sheet.merges = merges;

    let rowIdx = -1;
    for (const row of kids(kid(ws, 'sheetData'), 'row')) {
      rowIdx = row.hasAttribute('r') ? Number(row.getAttribute('r')) - 1 : rowIdx + 1;
      // A row height is in points; 15 pt is the default row, 20 px. Rows that Excel sized to their text
      // (no customHeight) are kept at that height too, so large text is not cut off.
      const rowStyle = ['1', 'true'].includes(row.getAttribute('customFormat') ?? '') ? xfs[Number(row.getAttribute('s') ?? 0)]?.st : undefined;
      if (rowStyle && rowIdx >= 0 && rowIdx <= 1_048_575) (sheet.rowStyles ??= {})[rowIdx] = rowStyle;
      const ht = Number(row.getAttribute('ht'));
      if (ht > 0 && rowIdx >= 0 && rowIdx <= 1_048_575) {
        const px = Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round((ht * 4) / 3)));
        if (px !== DEFAULT_ROW_HEIGHT) (sheet.rowHeights ??= {})[rowIdx] = px;
      }
      let colIdx = -1;
      for (const c of kids(row, 'c')) {
        const a = decodeAddr(c.getAttribute('r') ?? '');
        if (a) colIdx = a.c;
        else colIdx++;
        const r = a ? a.r : rowIdx;
        if (r > 1_048_575 || colIdx > 16_383) continue;
        const t = c.getAttribute('t');
        const f = kid(c, 'f')?.textContent;
        const v = kid(c, 'v')?.textContent ?? '';
        let raw: string;
        let text = false;
        if (f) raw = '=' + f;
        else if (t === 's') raw = shared[Number(v)] ?? '';
        else if (t === 'inlineStr') raw = richText(kid(c, 'is') ?? c);
        else if (t === 'b') raw = v === '1' ? 'TRUE' : 'FALSE';
        else raw = v;
        const xf = xfs[Number(c.getAttribute('s') ?? 0)];
        // A cell with no formatting of its own in a styled row or column is kept (with an empty style) even
        // when it is empty: Excel shows it unformatted, not with the row's or column's formatting.
        const inherited = sheet.rowStyles?.[r] ?? sheet.colStyles?.[colIdx];
        if (raw === '' && !xf?.st && !inherited) continue;
        if (!f && (t === 's' || t === 'inlineStr' || t === 'str') && typeof parseLiteral(raw) !== 'string') text = true;
        const cell: Cell = { raw };
        if (xf?.z) cell.z = xf.z;
        if (xf?.st) cell.st = xf.st;
        else if (inherited) cell.st = {};
        if (text) cell.text = true;
        (sheet.rows[r] ??= [])[colIdx] = cell;
      }
    }
    sheets.push(sheet);
  }
  if (!sheets.length) throw new Error('The workbook has no sheets');
  return { sheets, active: 0 };
}

// ---------- writing ----------

const esc = (s: string) =>
  s
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** The editor accepts ";" as an argument separator; the file must contain ",". */
const excelFormula = (f: string) =>
  f
    .split(/("(?:[^"]|"")*")/)
    .map((seg, i) => (i % 2 ? seg : seg.replace(/;/g, ',')))
    .join('');

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export async function writeWorkbook(book: Book): Promise<Uint8Array> {
  const ev = createEvaluator(book);
  const fmts: string[] = [];
  const fonts: string[] = ['<font><sz val="11"/><name val="Calibri"/></font>'];
  const fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders: string[] = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',
    '<border>' + ['left', 'right', 'top', 'bottom'].map((x) => `<${x} style="thin"><color rgb="FF000000"/></${x}>`).join('') + '<diagonal/></border>',
  ];
  const xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const xfIdx = new Map<string, number>();
  const reg = (list: string[], x: string) => {
    let i = list.indexOf(x);
    if (i < 0) i = list.push(x) - 1;
    return i;
  };
  const hex = (c?: string) => (isColor(c) ? c.slice(1).toUpperCase() : undefined);
  /** xf index for a cell's formatting (0 is the default style). */
  const styleIdx = (cell: Cell): number => {
    if (!cell.z && !cell.st) return 0;
    const key = JSON.stringify([cell.z, cell.st]);
    const hit = xfIdx.get(key);
    if (hit !== undefined) return hit;
    const st = cell.st ?? {};
    const sz = st.sz && st.sz >= 6 && st.sz <= 72 ? st.sz : 11;
    const col = hex(st.color);
    const font =
      st.b || st.i || st.u || sz !== 11 || col
        ? reg(fonts, `<font>${st.b ? '<b/>' : ''}${st.i ? '<i/>' : ''}${st.u ? '<u/>' : ''}<sz val="${sz}"/>${col ? `<color rgb="FF${col}"/>` : ''}<name val="Calibri"/></font>`)
        : 0;
    const bg = hex(st.bg);
    const fill = bg ? reg(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${bg}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const border = st.bd ? 1 : 0;
    const nf = cell.z ? reg(fmts, cell.z) + 164 : 0;
    const align = st.h === 'left' || st.h === 'center' || st.h === 'right' ? `<alignment horizontal="${st.h}"/>` : '';
    const xf = `<xf numFmtId="${nf}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${align ? ' applyAlignment="1">' + align + '</xf>' : '/>'}`;
    const i = xfs.push(xf) - 1;
    xfIdx.set(key, i);
    return i;
  };

  const sheetXml = book.sheets.map((s, si) => {
    const { cols } = usedSize(s);
    let out = `${HEAD}<worksheet ${NS}>`;
    const limit = Math.max(cols, 1) + 1000;
    const colIndexes = new Set<number>();
    for (const k of Object.keys(s.colWidths)) if (Number(k) < limit) colIndexes.add(Number(k));
    for (const k of Object.keys(s.colStyles ?? {})) if (Number(k) < limit) colIndexes.add(Number(k));
    if (colIndexes.size) {
      out += '<cols>';
      for (const k of [...colIndexes].sort((a, b) => a - b)) {
        const w = s.colWidths[k];
        const st = s.colStyles?.[k];
        const style = st ? styleIdx({ raw: '', st }) : 0;
        // A column that only has a style gets the editor's default width, so it looks the same here and there.
        const width = w !== undefined ? `width="${Math.max(0, (w - 5) / 7).toFixed(2)}" customWidth="1"` : 'width="9.57"';
        out += `<col min="${k + 1}" max="${k + 1}" ${width}${style ? ` style="${style}"` : ''}/>`;
      }
      out += '</cols>';
    }
    out += '<sheetData>';
    const heights = s.rowHeights ?? {};
    const rowIndexes = new Set<number>([...s.rows.keys()].filter((r) => s.rows[r]));
    for (const k of Object.keys(heights)) if (Number(k) <= 1_048_575) rowIndexes.add(Number(k));
    for (const k of Object.keys(s.rowStyles ?? {})) if (Number(k) <= 1_048_575) rowIndexes.add(Number(k));
    [...rowIndexes].sort((a, b) => a - b).forEach((r) => {
      const row = s.rows[r];
      let cells = '';
      row?.forEach((cell, c) => {
        if (!cell || (cell.raw === '' && !cell.st)) return;
        const ref = encodeAddr(r, c);
        const si2 = styleIdx(cell);
        const style = si2 ? ` s="${si2}"` : '';
        if (cell.raw === '') {
          cells += `<c r="${ref}"${style}/>`;
          return;
        }
        if (cell.raw[0] === '=' && !cell.text) {
          const v = ev.value(si, r, c);
          const f = `<f>${esc(excelFormula(cell.raw.slice(1)))}</f>`;
          if (isErr(v)) {
            const code = ERR_CODES.includes(v.err) ? v.err : '#VALUE!';
            cells += `<c r="${ref}"${style} t="e">${f}<v>${esc(code)}</v></c>`;
          } else if (typeof v === 'boolean') cells += `<c r="${ref}"${style} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
          else if (typeof v === 'string') cells += `<c r="${ref}"${style} t="str">${f}<v>${esc(v)}</v></c>`;
          else cells += `<c r="${ref}"${style}>${f}<v>${v ?? 0}</v></c>`;
          return;
        }
        const v = parseLiteral(cell.raw, cell.text);
        if (typeof v === 'number') cells += `<c r="${ref}"${style}><v>${v}</v></c>`;
        else if (typeof v === 'boolean') cells += `<c r="${ref}"${style} t="b"><v>${v ? 1 : 0}</v></c>`;
        else if (isErr(v)) cells += `<c r="${ref}"${style} t="e"><v>${esc(v.err)}</v></c>`;
        else cells += `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${esc(String(v))}</t></is></c>`;
      });
      const px = heights[r];
      const height = px ? ` ht="${((px * 3) / 4).toFixed(2)}" customHeight="1"` : '';
      const rs = s.rowStyles?.[r];
      const rowStyle = rs && styleIdx({ raw: '', st: rs });
      const format = rowStyle ? ` s="${rowStyle}" customFormat="1"` : '';
      if (cells || height || format) out += `<row r="${r + 1}"${format}${height}>${cells}</row>`;
    });
    out += '</sheetData>';
    if (s.merges?.length) {
      out += `<mergeCells count="${s.merges.length}">`;
      for (const m of s.merges) out += `<mergeCell ref="${encodeAddr(m.s.r, m.s.c)}:${encodeAddr(m.e.r, m.e.c)}"/>`;
      out += '</mergeCells>';
    }
    return out + '</worksheet>';
  });

  const enc = new TextEncoder();
  const files = new Map<string, Uint8Array>();
  const put = (n: string, s: string) => files.set(n, enc.encode(s));
  const n = book.sheets.length;

  put(
    '[Content_Types].xml',
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${book.sheets
      .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
      .join('')}</Types>`,
  );
  put('_rels/.rels', `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  put(
    'xl/workbook.xml',
    `${HEAD}<workbook ${NS} xmlns:r="${REL_NS}"><sheets>${book.sheets
      .map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets></workbook>`,
  );
  put(
    'xl/_rels/workbook.xml.rels',
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${book.sheets
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('')}<Relationship Id="rId${n + 1}" Type="${REL_NS}/styles" Target="styles.xml"/></Relationships>`,
  );
  // styles are written after the sheets, when all formats are known
  put(
    'xl/styles.xml',
    `${HEAD}<styleSheet ${NS}>${
      fmts.length ? `<numFmts count="${fmts.length}">${fmts.map((z, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(z)}"/>`).join('')}</numFmts>` : ''
    }<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs></styleSheet>`,
  );
  sheetXml.forEach((x, i) => put(`xl/worksheets/sheet${i + 1}.xml`, x));
  void encodeCol;
  void decodeCol;
  return zip(files);
}
