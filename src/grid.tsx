import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { KeyboardEvent, ClipboardEvent } from 'react';
import type { Sel, Sheet } from './lib/types';
import { selRect, selRects, usedSize } from './lib/types';
import type { Rect } from './lib/types';
import { display, isErr } from './lib/formula';
import type { Evaluator } from './lib/formula';
import { encodeCol } from './lib/addr';
import type { Dir } from './lib/ops';
import { DEFAULT_ROW_HEIGHT, MAX_ROW_HEIGHT, MIN_ROW_HEIGHT } from './lib/store';

/** Sizes at 100% zoom, Excel-like: 20 px rows, 72 px columns. */
const BASE_RH = 20;
const BASE_RHW = 40;
const BASE_W = 72;
const MIN_W = 12;

export interface Editing {
  r: number;
  c: number;
  value: string;
  src: 'cell' | 'bar';
}

interface Props {
  sheet: Sheet;
  si: number;
  ev: Evaluator;
  sel: Sel;
  onSelect: (s: Sel) => void;
  editing: Editing | null;
  onEditChange: (v: string) => void;
  onEditKey: (e: KeyboardEvent) => void;
  onStartEdit: () => void;
  onFill: (dir: Dir, count: number) => void;
  /** Called once, when the user lets go of the edge of a column header; undefined means the default width. */
  onColWidth: (c: number, w: number | undefined) => void;
  /** Called once, when the user lets go of the edge of a row header; undefined means the default height. */
  onRowHeight: (r: number, h: number | undefined) => void;
  onKeyDown: (e: KeyboardEvent) => void;
  onBeforeInput: (e: React.FormEvent<HTMLDivElement>) => void;
  onCopy: (e: ClipboardEvent, cut: boolean) => void;
  onPaste: (e: ClipboardEvent) => void;
  rootRef: React.RefObject<HTMLDivElement | null>;
  zoom: number;
  onZoom: (factor: number) => void;
}

export function Grid(p: Props) {
  const { sheet, sel, editing, zoom } = p;
  const RH = Math.max(6, Math.round(BASE_RH * zoom));
  const HH = RH;
  const RHW = Math.max(16, Math.round(BASE_RHW * zoom));
  // A column or row that is being resized: shown at its new size while the mouse is down, and only then
  // written to the workbook (one step of the history, and the document becomes changed).
  const [drag, setDrag] = useState<{ axis: 'col' | 'row'; index: number; size: number } | null>(null);
  const colW = (c: number) =>
    Math.max(2, Math.round(((drag?.axis === 'col' && drag.index === c ? drag.size : sheet.colWidths[c]) ?? BASE_W) * zoom));
  const rowH = (r: number) =>
    Math.max(2, Math.round(((drag?.axis === 'row' && drag.index === r ? drag.size : sheet.rowHeights?.[r]) ?? DEFAULT_ROW_HEIGHT) * zoom));
  const scroller = useRef<HTMLDivElement>(null);
  const content = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ x: 0, y: 0, w: 800, h: 600 });
  const fillRef = useRef<{ dir: Dir; count: number } | null>(null);
  const [fill, setFill] = useState<{ dir: Dir; count: number } | null>(null);

  const used = usedSize(sheet);
  const nRows = Math.max(100, used.rows + 30, sel.fr + 30, sel.ar + 30);
  const nCols = Math.max(26, used.cols + 5, sel.fc + 3, sel.ac + 3);

  const offs = useMemo(() => {
    const o = new Array<number>(nCols + 1);
    o[0] = 0;
    for (let c = 0; c < nCols; c++) o[c + 1] = o[c] + colW(c);
    return o;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet.colWidths, nCols, zoom, drag]);
  const roffs = useMemo(() => {
    const o = new Array<number>(nRows + 1);
    o[0] = 0;
    for (let r = 0; r < nRows; r++) o[r + 1] = o[r] + rowH(r);
    return o;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet.rowHeights, nRows, zoom, drag]);
  const totalW = offs[nCols];
  const totalH = roffs[nRows];

  useEffect(() => {
    const el = scroller.current;
    if (!el) return;
    const upd = () => setScroll({ x: el.scrollLeft, y: el.scrollTop, w: el.clientWidth, h: el.clientHeight });
    upd();
    const ro = new ResizeObserver(upd);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // scroll to the moving cell
  useEffect(() => {
    const el = scroller.current;
    if (!el || sel.noScroll) return;
    const top = roffs[sel.fr] ?? 0;
    const bottom = roffs[sel.fr + 1] ?? top;
    if (top < el.scrollTop) el.scrollTop = top;
    else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
    const left = offs[sel.fc] ?? 0;
    const right = offs[sel.fc + 1] ?? left;
    if (left < el.scrollLeft) el.scrollLeft = left;
    else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth;
  }, [sel.fr, sel.fc, sel.noScroll, offs, roffs]);

  const prevZoom = useRef(zoom);
  useLayoutEffect(() => {
    const el = scroller.current;
    if (el && prevZoom.current !== zoom) {
      const k = zoom / prevZoom.current;
      el.scrollTop *= k;
      el.scrollLeft *= k;
      setScroll((s) => ({ ...s, x: el.scrollLeft, y: el.scrollTop }));
    }
    prevZoom.current = zoom;
  }, [zoom]);

  const onZoomRef = useRef(p.onZoom);
  onZoomRef.current = p.onZoom;
  useEffect(() => {
    const el = p.rootRef.current;
    if (!el) return;
    const wheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      onZoomRef.current(Math.exp(-e.deltaY * 0.002));
    };
    el.addEventListener('wheel', wheel, { passive: false });
    return () => el.removeEventListener('wheel', wheel);
  }, [p.rootRef]);

  const colAt = (x: number) => {
    let lo = 0;
    let hi = nCols - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (offs[mid] <= x) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  const rowAt = (y: number) => {
    let lo = 0;
    let hi = nRows - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (roffs[mid] <= y) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };
  const cellAt = (e: { clientX: number; clientY: number }) => {
    const rect = content.current!.getBoundingClientRect();
    return {
      r: rowAt(Math.max(0, e.clientY - rect.top)),
      c: colAt(Math.max(0, e.clientX - rect.left)),
    };
  };

  const onMouseDown = (e: React.MouseEvent) => {
    if (e.button !== 0 || (e.target as HTMLElement).tagName === 'INPUT') return;
    e.preventDefault();
    p.rootRef.current?.focus();
    const start = cellAt(e);
    const additive = e.ctrlKey || e.metaKey;
    const extra = additive ? selRects(sel) : e.shiftKey ? sel.extra : undefined;
    if (e.shiftKey && !additive) p.onSelect({ ...sel, fr: start.r, fc: start.c });
    else p.onSelect({ ar: start.r, ac: start.c, fr: start.r, fc: start.c, extra });
    const move = (ev: MouseEvent) => {
      const t = cellAt(ev);
      const shift = e.shiftKey && !additive;
      p.onSelect({ ar: shift ? sel.ar : start.r, ac: shift ? sel.ac : start.c, fr: t.r, fc: t.c, extra });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  /** Resizing a column or a row by its header edge: live while the mouse is down, one change when it is let go. */
  const startResize = (axis: 'col' | 'row', index: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const start = axis === 'col' ? e.clientX : e.clientY;
    const size0 = (axis === 'col' ? sheet.colWidths[index] : sheet.rowHeights?.[index]) ?? (axis === 'col' ? BASE_W : DEFAULT_ROW_HEIGHT);
    const [min, max] = axis === 'col' ? [MIN_W, 100_000] : [MIN_ROW_HEIGHT, MAX_ROW_HEIGHT];
    let size = size0;
    const move = (ev: MouseEvent) => {
      size = Math.min(max, Math.max(min, Math.round(size0 + ((axis === 'col' ? ev.clientX : ev.clientY) - start) / zoom)));
      setDrag({ axis, index, size });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (size === size0) return;
      if (axis === 'col') p.onColWidth(index, size);
      else p.onRowHeight(index, size);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  /** Double-clicking the edge puts a column or row back to its default size. */
  const resetSize = (axis: 'col' | 'row', index: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (axis === 'col') p.onColWidth(index, undefined);
    else p.onRowHeight(index, undefined);
  };

  const { r1, r2, c1, c2 } = selRect(sel);
  const rStart = Math.max(0, rowAt(scroll.y) - 3);
  const rEnd = Math.min(nRows, rowAt(scroll.y + scroll.h) + 4);
  const cStart = Math.max(0, colAt(scroll.x) - 1);
  const cEnd = Math.min(nCols, colAt(scroll.x + scroll.w) + 2);
  const rowsInView = Array.from({ length: Math.max(0, rEnd - rStart) }, (_, i) => rStart + i);
  const colsInView = Array.from({ length: Math.max(0, cEnd - cStart) }, (_, i) => cStart + i);

  const cells = rowsInView.flatMap((r) => {
    const row = sheet.rows[r];
    if (!row) return [];
    return colsInView.flatMap((c) => {
      const cell = row[c];
      if (!cell || (cell.raw === '' && !cell.st)) return [];
      if (editing && editing.r === r && editing.c === c && editing.src === 'cell') return [];
      const v = p.ev.value(p.si, r, c);
      const cls = typeof v === 'number' ? 'xe-cell xe-num' : isErr(v) ? 'xe-cell xe-err' : typeof v === 'boolean' ? 'xe-cell xe-ctr' : 'xe-cell';
      const st = cell.st;
      const style: React.CSSProperties = { left: offs[c], top: roffs[r], width: offs[c + 1] - offs[c], height: roffs[r + 1] - roffs[r], lineHeight: roffs[r + 1] - roffs[r] + 'px' };
      if (st) {
        if (st.b) style.fontWeight = 700;
        if (st.i) style.fontStyle = 'italic';
        if (st.u) style.textDecoration = 'underline';
        if (st.sz) style.fontSize = (st.sz * 14 * zoom) / 11;
        if (st.color) style.color = st.color;
        if (st.bg) style.background = st.bg;
        if (st.h) style.textAlign = st.h;
        if (st.bd) style.boxShadow = 'inset 0 0 0 1px #000';
      }
      // A cell that exists is shown with its own formatting only: in a filled column or row it covers the fill.
      if (!style.background && (sheet.rowStyles?.[r]?.bg || sheet.colStyles?.[c]?.bg)) style.background = 'var(--bg)';
      return [
        <div key={r * 20000 + c} className={cls} style={style}>
          {display(v, cell.z)}
        </div>,
      ];
    });
  });

  const rects = selRects(sel);
  const colOn = (c: number) => rects.some((x) => c >= x.c1 && c <= x.c2);
  const rowOn = (r: number) => rects.some((x) => r >= x.r1 && r <= x.r2);

  /** Column/row selection via headers: click, drag, Shift extends, Ctrl adds. */
  const startHeader = (kind: 'col' | 'row', idx: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;
    e.preventDefault();
    p.rootRef.current?.focus();
    const additive = e.ctrlKey || e.metaKey;
    const extra: Rect[] | undefined = additive ? rects : e.shiftKey ? sel.extra : undefined;
    const whole = (a: number, b: number): Sel =>
      kind === 'col'
        ? { ar: 0, ac: a, fr: nRows - 1, fc: b, noScroll: true, extra, whole: 'col' }
        : { ar: a, ac: 0, fr: b, fc: nCols - 1, noScroll: true, extra, whole: 'row' };
    const anchor = e.shiftKey && !additive ? (kind === 'col' ? sel.ac : sel.ar) : idx;
    p.onSelect(whole(anchor, idx));
    const move = (ev: MouseEvent) => {
      const t = cellAt(ev);
      p.onSelect(whole(anchor, kind === 'col' ? t.c : t.r));
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const colHeads = colsInView.map((c) => (
      <div key={c} className={'xe-hd' + (colOn(c) ? ' xe-on' : '')} style={{ left: offs[c], width: offs[c + 1] - offs[c] }} onMouseDown={(e) => startHeader('col', c, e)}>
        {encodeCol(c)}
        <span className="xe-grip" onMouseDown={(e) => startResize('col', c, e)} onDoubleClick={(e) => resetSize('col', c, e)} />
      </div>
  ));
  const rowHeads = rowsInView.map((r) => (
      <div key={r} className={'xe-hd' + (rowOn(r) ? ' xe-on' : '')} style={{ top: roffs[r], height: roffs[r + 1] - roffs[r], width: RHW }} onMouseDown={(e) => startHeader('row', r, e)}>
        {r + 1}
        <span className="xe-rgrip" onMouseDown={(e) => startResize('row', r, e)} onDoubleClick={(e) => resetSize('row', r, e)} />
      </div>
  ));

  // The fill of a whole column or row shows in the cells that do not exist; a row's wins over a column's.
  const fills = [
    ...colsInView.flatMap((c) => {
      const bg = sheet.colStyles?.[c]?.bg;
      return bg ? [<div key={'c' + c} className="xe-fill" style={{ left: offs[c], top: 0, width: offs[c + 1] - offs[c], height: totalH, background: bg }} />] : [];
    }),
    ...rowsInView.flatMap((r) => {
      const bg = sheet.rowStyles?.[r]?.bg;
      return bg ? [<div key={'r' + r} className="xe-fill" style={{ left: 0, top: roffs[r], width: totalW, height: roffs[r + 1] - roffs[r], background: bg }} />] : [];
    }),
  ];

  // grid lines
  const vlines = colsInView.map((c) => <div key={c} className="xe-vline" style={{ left: offs[c + 1] - 1, height: totalH }} />);
  const hlines = rowsInView.map((r) => <div key={r} className="xe-hline" style={{ top: roffs[r + 1] - 1, width: totalW }} />);

  const startFill = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const move = (ev: MouseEvent) => {
      const t = cellAt(ev);
      const dr = t.r > r2 ? t.r - r2 : t.r < r1 ? t.r - r1 : 0;
      const dc = t.c > c2 ? t.c - c2 : t.c < c1 ? t.c - c1 : 0;
      let f: { dir: Dir; count: number } | null = null;
      if (Math.abs(dr) >= Math.abs(dc) && dr !== 0) f = { dir: dr > 0 ? 'down' : 'up', count: Math.abs(dr) };
      else if (dc !== 0) f = { dir: dc > 0 ? 'right' : 'left', count: Math.abs(dc) };
      if (f?.dir === 'up') f.count = Math.min(f.count, r1);
      if (f?.dir === 'left') f.count = Math.min(f.count, c1);
      if (f && f.count <= 0) f = null;
      fillRef.current = f;
      setFill(f);
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      const f = fillRef.current;
      fillRef.current = null;
      setFill(null);
      if (f) p.onFill(f.dir, f.count);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const box = (x: Rect) => ({ left: offs[x.c1], top: roffs[x.r1], width: offs[x.c2 + 1] - offs[x.c1], height: roffs[x.r2 + 1] - roffs[x.r1] });
  const rect = box({ r1, c1, r2, c2 });
  const act = { left: offs[sel.ac], top: roffs[sel.ar], width: offs[sel.ac + 1] - offs[sel.ac], height: roffs[sel.ar + 1] - roffs[sel.ar] };

  return (
    <div
      className="xe-grid"
      style={{ '--rh': RH + 'px', '--fs': 14 * zoom + 'px', '--hfs': 13 * zoom + 'px' } as React.CSSProperties}
      ref={p.rootRef}
      tabIndex={0}
      onKeyDown={p.onKeyDown}
      onBeforeInput={p.onBeforeInput}
      onCopy={(e) => p.onCopy(e, false)}
      onCut={(e) => p.onCopy(e, true)}
      onPaste={p.onPaste}
    >
      <div className="xe-corner" style={{ width: RHW, height: HH }} />
      <div className="xe-colheads" style={{ left: RHW, height: HH }}>
        <div style={{ transform: `translateX(${-scroll.x}px)`, width: totalW, height: HH, position: 'relative' }}>{colHeads}</div>
      </div>
      <div className="xe-rowheads" style={{ top: HH, width: RHW }}>
        <div style={{ transform: `translateY(${-scroll.y}px)`, height: totalH, position: 'relative' }}>{rowHeads}</div>
      </div>
      <div
        className="xe-scroller"
        ref={scroller}
        style={{ left: RHW, top: HH }}
        onScroll={(e) => {
          const el = e.currentTarget;
          setScroll((s) => ({ ...s, x: el.scrollLeft, y: el.scrollTop }));
        }}
      >
        <div ref={content} className="xe-content" style={{ width: totalW, height: totalH }} onMouseDown={onMouseDown} onDoubleClick={p.onStartEdit}>
          {hlines}
          {vlines}
          {fills}
          {cells}
          {rects.map((x, i) => (
            <div key={i} className="xe-selrect" style={box(x)} />
          ))}
          {fill && (
            <div
              className="xe-fillrect"
              style={{
                left: offs[fill.dir === 'left' ? c1 - fill.count : c1],
                top: roffs[fill.dir === 'up' ? r1 - fill.count : r1],
                width: offs[fill.dir === 'right' ? c2 + 1 + fill.count : c2 + 1] - offs[fill.dir === 'left' ? c1 - fill.count : c1],
                height: roffs[(fill.dir === 'down' ? r2 + fill.count : r2) + 1] - roffs[fill.dir === 'up' ? r1 - fill.count : r1],
              }}
            />
          )}
          <div className="xe-active" style={act} />
          {!editing && <div className="xe-handle" style={{ left: rect.left + rect.width - 5, top: rect.top + rect.height - 5 }} onMouseDown={startFill} />}
          {editing && editing.src === 'cell' && (
            <input
              className="xe-editor"
              autoFocus
              spellCheck={false}
              style={{ left: offs[editing.c], top: roffs[editing.r], width: offs[editing.c + 1] - offs[editing.c], height: roffs[editing.r + 1] - roffs[editing.r] }}
              value={editing.value}
              onChange={(e) => p.onEditChange(e.target.value)}
              onKeyDown={(e) => {
                e.stopPropagation();
                p.onEditKey(e);
              }}
              onFocus={(e) => e.currentTarget.setSelectionRange(e.currentTarget.value.length, e.currentTarget.value.length)}
            />
          )}
        </div>
      </div>
    </div>
  );
}
