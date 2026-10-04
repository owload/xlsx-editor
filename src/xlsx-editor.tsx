import { useEffect, useImperativeHandle, useMemo, useReducer, useRef, useState } from 'react';
import type { ClipboardEvent, KeyboardEvent, Ref } from 'react';
import { Grid } from './grid';
import type { Editing } from './grid';
import { Ribbon } from './ribbon';
import { FindBar } from './find-bar';
import { SheetTabs } from './sheet-tabs';
import { autosum, changeDecimals, clearContents, fillRange, findNext, norm, patchStyle, replaceAll, replaceAt, setFormat, sortRows } from './lib/ops';
import type { Dir } from './lib/ops';
import type { Book, Cell, Sel, Style } from './lib/types';
import { getCell, inheritedStyle, intervals, selRect, selRects, usedSize } from './lib/types';
import { collectUsedColors } from './lib/colors';
import { createEvaluator, display, isErr, parseLiteral } from './lib/formula';
import {
  addSheet,
  deleteLines,
  deleteSheet,
  initialHistory,
  insertLines,
  isDirty,
  moveSheet,
  newBook,
  reducer,
  renameSheet,
  setCells,
  setColWidth,
  setRowHeight,
} from './lib/store';
import { readWorkbook, writeWorkbook } from './lib/xlsx-io';
import { encodeAddr } from './lib/addr';
import { parseTSV, toTSV } from './lib/clipboard';
import './xlsx-editor.css';

const DEFAULT_MAX_FILE_BYTES = 100 * 1024 * 1024;
const MAX_CELL_CHARS = 32767;
const MAX_R = 1_048_575;
const MAX_C = 16_383;

/** Imperative API, available through the `ref` prop. */
export interface XlsxEditorHandle {
  /** Serializes the workbook and passes it to `onSave`. Resolves once `onSave` has finished. */
  save(): Promise<void>;
  /** True if there are changes that have not been passed to `onSave` yet. */
  isDirty(): boolean;
}

export interface XlsxEditorProps {
  /**
   * Contents of the .xlsx file (typically decrypted by the host). Empty or null starts a blank
   * workbook. It is read once on mount; to open another file, remount the component with a new `key`.
   * The editor never keeps a reference to this buffer after parsing.
   */
  data: Uint8Array | ArrayBuffer | null | undefined;
  /** Shown in the title bar only; the editor never reads or writes files by name. */
  fileName: string;
  /**
   * Called with the serialized .xlsx bytes when the user saves (Save button or Ctrl+S).
   * If it throws or rejects, the error is shown and the document stays dirty.
   */
  onSave: (data: Uint8Array) => void | Promise<void>;
  /** Called whenever the "has unsaved changes" state flips. */
  onDirtyChange?: (dirty: boolean) => void;
  /** Called when loading or saving fails (the message is also shown in the UI). */
  onError?: (error: Error) => void;
  /**
   * Called when the user clicks the close button in the title bar. The button is shown only if this is
   * given. The editor closes nothing itself and does not ask about unsaved changes: the host does.
   */
  onClose?: () => void;
  /** Reject files larger than this many bytes. Default: 100 MiB. */
  maxFileBytes?: number;
  /**
   * Keep copied cells inside the editor only: nothing is written to, or read from, the system
   * clipboard. Recommended when plaintext must not leave the encrypted storage. Default: false.
   */
  internalClipboardOnly?: boolean;
  className?: string;
  ref?: Ref<XlsxEditorHandle>;
}

const editText = (cell?: Cell) => {
  if (!cell) return '';
  return cell.text && typeof parseLiteral(cell.raw) !== 'string' ? "'" + cell.raw : cell.raw;
};

function toCell(value: string, prev?: Cell, inherited?: Style): Cell | undefined {
  const v = value.slice(0, MAX_CELL_CHARS);
  // A cell that does not exist starts with the formatting of its row or column; one that exists keeps its own
  // (an empty style, if it has none in a styled row or column, so that their formatting does not appear in it).
  const st = prev ? (prev.st ?? (inherited ? {} : undefined)) : inherited;
  const keep = { ...(prev?.z && { z: prev.z }), ...(st && { st }) };
  if (v[0] === "'") return { raw: v.slice(1), text: true, ...keep };
  return norm({ raw: v, ...keep });
}

const clamp = (n: number, max: number) => Math.max(0, Math.min(max, n));

type LoadState = { status: 'loading' } | { status: 'ready' } | { status: 'error'; message: string };

const toError = (e: unknown, fallback: string) => (e instanceof Error ? e : new Error(fallback));

export function XlsxEditor({
  data,
  fileName,
  onSave,
  onDirtyChange,
  onError,
  onClose,
  maxFileBytes = DEFAULT_MAX_FILE_BYTES,
  internalClipboardOnly = false,
  className,
  ref,
}: XlsxEditorProps) {
  const [h, dispatch] = useReducer(reducer, undefined, () => initialHistory(newBook()));
  const [load, setLoad] = useState<LoadState>({ status: 'loading' });
  const [sel, setSel] = useState<Sel>({ ar: 0, ac: 0, fr: 0, fc: 0 });
  const [editing, setEditing] = useState<Editing | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const clip = useRef<{ text: string; cells: (Cell | undefined)[][] } | null>(null);
  const [zoom, setZoom] = useState(1);
  const [find, setFind] = useState<{ replace: boolean; message: string } | null>(null);
  // What saving this file would drop (the editor rebuilds it from what it supports); empty if nothing.
  const [losses, setLosses] = useState<string[]>([]);

  const { book } = h;
  const si = book.active;
  const sheet = book.sheets[si];
  const ev = useMemo(() => createEvaluator(book), [book]);
  const usedColors = useMemo(() => collectUsedColors(book), [book]);
  const dirty = isDirty(h);

  const fail = (e: unknown, fallback: string) => {
    const err = toError(e, fallback);
    setError(err.message);
    onError?.(err);
  };

  // The input buffer is only read once, on mount (see XlsxEditorProps.data).
  const initialData = useRef(data);
  const callbacks = useRef({ onError, onDirtyChange });
  callbacks.current = { onError, onDirtyChange };
  useEffect(() => {
    let cancelled = false;
    const input = initialData.current;
    initialData.current = undefined;
    (async () => {
      try {
        const bytes = input instanceof Uint8Array ? input : input ? new Uint8Array(input) : null;
        if (bytes && bytes.byteLength > maxFileBytes) throw new Error('The file is too large to open');
        const loaded = bytes && bytes.byteLength > 0 ? await readWorkbook(bytes) : newBook();
        if (cancelled) return;
        dispatch({ type: 'load', book: loaded });
        setLoad({ status: 'ready' });
        if (bytes && bytes.byteLength > 0) {
          // Looked at in the background; the file was read above, so a failure here only means no note.
          import('./lib/inspect')
            .then((m) => m.inspectWorkbook(bytes))
            .then((r) => { if (!cancelled) setLosses(r.unsupported.map((u) => u.label)); })
            .catch(() => undefined);
        }
      } catch (e) {
        if (cancelled) return;
        const err = toError(e, 'Failed to open the file');
        setLoad({ status: 'error', message: err.message });
        callbacks.current.onError?.(err);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [maxFileBytes]);

  useEffect(() => {
    callbacks.current.onDirtyChange?.(dirty);
  }, [dirty]);

  const savingRef = useRef(false);
  const save = async () => {
    if (savingRef.current || load.status !== 'ready') return;
    savingRef.current = true;
    setSaving(true);
    try {
      // Include a cell that is still being edited.
      let toWrite = book;
      let rev = h.rev;
      if (editing) {
        const prev = getCell(sheet, editing.r, editing.c);
        if (editing.value !== editText(prev)) {
          toWrite = setCells(book, si, [{ r: editing.r, c: editing.c, cell: toCell(editing.value, prev, inheritedStyle(sheet, editing.r, editing.c)) }]);
          rev += 1;
        }
        applyEdit(editing);
        setEditing(null);
      }
      const bytes = await writeWorkbook(toWrite);
      await onSave(bytes);
      dispatch({ type: 'saved', rev });
      setError(null);
      setLosses([]); // the file has been rewritten: what it could not hold is gone
    } catch (e) {
      fail(e, 'Failed to save');
    } finally {
      savingRef.current = false;
      setSaving(false);
    }
  };

  const latest = useRef({ save, dirty });
  latest.current = { save, dirty };
  useImperativeHandle(ref, () => ({ save: () => latest.current.save(), isDirty: () => latest.current.dirty }), []);

  // ---------- editing ----------

  const applyEdit = (ed: Editing) => {
    const prev = getCell(sheet, ed.r, ed.c);
    if (ed.value === editText(prev)) return;
    dispatch({ type: 'commit', fn: (b) => setCells(b, si, [{ r: ed.r, c: ed.c, cell: toCell(ed.value, prev, inheritedStyle(sheet, ed.r, ed.c)) }]) });
  };
  const flush = () => {
    if (!editing) return;
    applyEdit(editing);
    setEditing(null);
  };
  const act = (fn: (b: Book) => Book) => {
    flush();
    dispatch({ type: 'commit', fn });
    rootRef.current?.focus();
  };

  const startEdit = (initialValue?: string, src: Editing['src'] = 'cell') => {
    if (editing) return;
    const value = initialValue ?? editText(getCell(sheet, sel.ar, sel.ac));
    setEditing({ r: sel.ar, c: sel.ac, value, src });
  };

  const commit = (dr: number, dc: number) => {
    if (!editing) return;
    applyEdit(editing);
    setEditing(null);
    const r = clamp(editing.r + dr, MAX_R);
    const c = clamp(editing.c + dc, MAX_C);
    setSel({ ar: r, ac: c, fr: r, fc: c });
    rootRef.current?.focus();
  };

  const onEditKey = (e: KeyboardEvent) => {
    // Ctrl/Cmd+S also works while a cell or the formula bar is being edited (the edit is saved too).
    if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      void save();
      return;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      commit(e.shiftKey ? -1 : 1, 0);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commit(0, e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setEditing(null);
      rootRef.current?.focus();
    }
  };

  const select = (s: Sel) => {
    flush();
    setSel(s);
  };

  // ---------- selection and clipboard ----------

  const used = usedSize(sheet);
  const dataRect = () => {
    const { r1, r2, c1, c2 } = selRect(sel);
    return { r1, c1, r2: Math.max(r1, Math.min(r2, used.rows - 1)), c2: Math.max(c1, Math.min(c2, used.cols - 1)) };
  };

  const clearSel = () =>
    dispatch({ type: 'commit', fn: (b) => selRects(sel).reduce((bk, rc) => clearContents(bk, si, { ...rc, r2: Math.min(rc.r2, Math.max(used.rows - 1, rc.r1)), c2: Math.min(rc.c2, Math.max(used.cols - 1, rc.c1)) }), b) });

  const onCopy = (e: ClipboardEvent, cut: boolean) => {
    if (editing) return;
    const { r1, r2, c1, c2 } = dataRect();
    const text: string[][] = [];
    const cells: (Cell | undefined)[][] = [];
    for (let r = r1; r <= r2; r++) {
      text.push([]);
      cells.push([]);
      for (let c = c1; c <= c2; c++) {
        const cell = getCell(sheet, r, c);
        cells[r - r1].push(cell);
        text[r - r1].push(display(ev.value(si, r, c), cell?.z));
      }
    }
    const tsv = toTSV(text);
    clip.current = { text: tsv, cells };
    // In internal-only mode the system clipboard is deliberately left empty.
    e.clipboardData.setData('text/plain', internalClipboardOnly ? '' : tsv);
    e.preventDefault();
    if (cut) clearSel();
  };

  const onPaste = (e: ClipboardEvent) => {
    if (editing) return;
    e.preventDefault();
    const text = internalClipboardOnly ? '' : e.clipboardData.getData('text/plain');
    const { r1, c1 } = selRect(sel);
    let grid: (Cell | undefined)[][];
    if (internalClipboardOnly) {
      if (!clip.current) return;
      grid = clip.current.cells;
    } else {
      if (!text) return;
      // Our own copy keeps formulas and formatting; anything else is parsed as plain TSV.
      const internal = clip.current && clip.current.text === toTSV(parseTSV(text)) ? clip.current.cells : null;
      grid = internal ?? parseTSV(text).map((row) => row.map((v) => toCell(v)));
    }
    const width = Math.max(...grid.map((r) => r.length), 1);
    if (grid.length * width > 1_000_000) return setError('The pasted range is too large');
    const edits: { r: number; c: number; cell?: Cell }[] = [];
    grid.forEach((row, i) =>
      row.forEach((cell, j) => {
        if (r1 + i <= MAX_R && c1 + j <= MAX_C) edits.push({ r: r1 + i, c: c1 + j, cell });
      }),
    );
    dispatch({ type: 'commit', fn: (b) => setCells(b, si, edits) });
    setSel({ ar: r1, ac: c1, fr: Math.min(MAX_R, r1 + grid.length - 1), fc: Math.min(MAX_C, c1 + width - 1) });
  };

  const onKeyDown = (e: KeyboardEvent) => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (mod && (k === 'z' || k === 'Z')) {
      e.preventDefault();
      dispatch({ type: e.shiftKey ? 'redo' : 'undo' });
      return;
    }
    if (mod && (k === 'y' || k === 'Y')) {
      e.preventDefault();
      dispatch({ type: 'redo' });
      return;
    }
    if (mod && (k === 's' || k === 'S')) {
      e.preventDefault();
      void save();
      return;
    }
    if (mod && /^[biuBIU]$/.test(k)) {
      e.preventDefault();
      const key = k.toLowerCase() as 'b' | 'i' | 'u';
      return patch({ [key]: !activeCell?.st?.[key] });
    }
    if (mod && /^[fFhH]$/.test(k)) {
      e.preventDefault();
      return setFind({ replace: /h/i.test(k), message: '' });
    }
    if (mod && /^[dDrR]$/.test(k)) {
      e.preventDefault();
      const rc = selRect(sel);
      const down = /d/i.test(k);
      if (down ? rc.r2 > rc.r1 : rc.c2 > rc.c1) {
        const src = down ? { ...rc, r2: rc.r1 } : { ...rc, c2: rc.c1 };
        act((b) => fillRange(b, si, src, down ? 'down' : 'right', down ? rc.r2 - rc.r1 : rc.c2 - rc.c1));
      }
      return;
    }
    if (mod && (k === 'a' || k === 'A')) {
      e.preventDefault();
      setSel({ ar: 0, ac: 0, fr: Math.max(0, used.rows - 1), fc: Math.max(0, used.cols - 1), noScroll: true });
      return;
    }
    const move = (dr: number, dc: number) => {
      e.preventDefault();
      if (e.shiftKey) setSel({ ...sel, fr: clamp(sel.fr + dr, MAX_R), fc: clamp(sel.fc + dc, MAX_C), noScroll: false });
      else {
        const r = clamp(sel.ar + dr, MAX_R);
        const c = clamp(sel.ac + dc, MAX_C);
        setSel({ ar: r, ac: c, fr: r, fc: c });
      }
    };
    if (k === 'ArrowUp') return move(-1, 0);
    if (k === 'ArrowDown') return move(1, 0);
    if (k === 'ArrowLeft') return move(0, -1);
    if (k === 'ArrowRight') return move(0, 1);
    if (k === 'PageUp') return move(-20, 0);
    if (k === 'PageDown') return move(20, 0);
    if (k === 'Tab') return move(0, e.shiftKey ? -1 : 1);
    if (k === 'Home') {
      e.preventDefault();
      return setSel({ ar: sel.ar, ac: 0, fr: sel.ar, fc: 0 });
    }
    if (k === 'Enter') {
      e.preventDefault();
      return e.shiftKey ? move(-1, 0) : startEdit();
    }
    if (k === 'F2') return startEdit();
    if (k === 'Delete' || k === 'Backspace') {
      e.preventDefault();
      return clearSel();
    }
    if (k.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      startEdit(k);
    }
  };

  // ---------- structure ----------

  const stats = useMemo(() => {
    const { r1, r2, c1, c2 } = selRect(sel);
    if ((r2 - r1 + 1) * (c2 - c1 + 1) < 2) return null;
    let sum = 0;
    let n = 0;
    const rEnd = Math.min(r2, used.rows - 1);
    const cEnd = Math.min(c2, used.cols - 1);
    for (let r = r1; r <= rEnd; r++)
      for (let c = c1; c <= cEnd; c++) {
        const v = ev.value(si, r, c);
        if (typeof v === 'number') {
          sum += v;
          n++;
        } else if (isErr(v)) return null;
      }
    return n ? { sum, n, avg: sum / n } : null;
  }, [sel, ev, si, used.rows, used.cols]);

  const activeCell = getCell(sheet, sel.ar, sel.ac);
  const barValue = editing ? editing.value : editText(activeCell);
  const patch = (p: Partial<Style>) => act((b) => selRects(sel).reduce((bk, rc) => patchStyle(bk, si, rc, p), b));
  const lines = (axis: 'r' | 'c', below: boolean | null) =>
    act((b) => {
      let out = b;
      for (const [a, z] of intervals(selRects(sel), axis)) {
        out = below === null ? deleteLines(out, si, axis, a, z - a + 1) : insertLines(out, si, axis, below ? z + 1 : a, z - a + 1);
      }
      return out;
    });
  const onFill = (dir: Dir, count: number) => {
    const rc = selRect(sel);
    act((b) => fillRange(b, si, rc, dir, count));
    const ext = { ...rc };
    if (dir === 'down') ext.r2 += count;
    else if (dir === 'up') ext.r1 -= count;
    else if (dir === 'right') ext.c2 += count;
    else ext.c1 -= count;
    setSel({ ar: ext.r1, ac: ext.c1, fr: ext.r2, fc: ext.c2 });
  };

  if (load.status === 'loading') {
    return (
      <div className={'xe' + (className ? ' ' + className : '')}>
        <div className="xe-state">Loading…</div>
      </div>
    );
  }
  if (load.status === 'error') {
    return (
      <div className={'xe' + (className ? ' ' + className : '')}>
        <div className="xe-state xe-error" role="alert">
          {load.message}
        </div>
      </div>
    );
  }

  return (
    <div className={'xe' + (className ? ' ' + className : '')}>
      <Ribbon
        st={activeCell?.st}
        z={activeCell?.z}
        canUndo={!!h.past.length}
        canRedo={!!h.future.length}
        fileName={fileName}
        dirty={dirty}
        saving={saving}
        usedColors={usedColors}
        onClose={onClose}
        a={{
          save: () => void save(),
          undo: () => dispatch({ type: 'undo' }),
          redo: () => dispatch({ type: 'redo' }),
          patch,
          setFormat: (z) => act((b) => selRects(sel).reduce((bk, rc) => setFormat(bk, si, rc, z), b)),
          decimals: (d) => act((b) => selRects(sel).reduce((bk, rc) => setFormat(bk, si, rc, changeDecimals(activeCell?.z, d)), b)),
          autosum: () => act((b) => autosum(b, si, sel.ar, sel.ac) ?? b),
          sort: (asc) => {
            const rc = selRect(sel);
            const region = rc.r2 > rc.r1 ? rc : { r1: 0, c1: 0, r2: used.rows - 1, c2: used.cols - 1 };
            act((b) => sortRows(b, si, ev, region, sel.ac, asc));
          },
          find: (replace) => setFind({ replace, message: '' }),
          insertRows: (below) => lines('r', below),
          deleteRows: () => lines('r', null),
          insertCols: (right) => lines('c', right),
          deleteCols: () => lines('c', null),
          clear: () => act((b) => selRects(sel).reduce((bk, rc) => clearContents(bk, si, { ...rc, r2: Math.min(rc.r2, Math.max(used.rows - 1, rc.r1)), c2: Math.min(rc.c2, Math.max(used.cols - 1, rc.c1)) }), b)),
        }}
      />
      {find && (
        <FindBar
          key={String(find.replace)}
          replace={find.replace}
          message={find.message}
          onClose={() => {
            setFind(null);
            rootRef.current?.focus();
          }}
          onNext={(q, cs) => {
            const hit = findNext(book, si, ev, q, cs, { r: sel.ar, c: sel.ac });
            if (hit) {
              setSel({ ar: hit.r, ac: hit.c, fr: hit.r, fc: hit.c });
              setFind({ ...find, message: '' });
            } else setFind({ ...find, message: 'Not found' });
          }}
          onReplace={(q, to, cs) => {
            dispatch({ type: 'commit', fn: (b) => replaceAt(b, si, sel.ar, sel.ac, q, to, cs) });
            const hit = findNext(book, si, ev, q, cs, { r: sel.ar, c: sel.ac });
            if (hit) setSel({ ar: hit.r, ac: hit.c, fr: hit.r, fc: hit.c });
          }}
          onReplaceAll={(q, to, cs) => {
            const res = replaceAll(book, si, q, to, cs);
            if (res.count) dispatch({ type: 'commit', fn: (b) => replaceAll(b, si, q, to, cs).book });
            setFind({ ...find, message: `Replaced: ${res.count}` });
          }}
        />
      )}

      {error && (
        <div className="xe-error" role="alert" onClick={() => setError(null)}>
          {error}
        </div>
      )}

      {losses.length > 0 && (
        <div className="xe-note" role="note">
          <span>Saving this file here will drop: {losses.map((l) => l.toLowerCase()).join(', ')}.</span>
          <button onClick={() => setLosses([])}>Got it</button>
        </div>
      )}

      <div className="xe-formulabar">
        <div className="xe-namebox">{encodeAddr(sel.ar, sel.ac)}</div>
        <span className="xe-fx">fx</span>
        <input
          value={barValue}
          spellCheck={false}
          onFocus={() => startEdit(undefined, 'bar')}
          onChange={(e) => {
            if (editing) setEditing({ ...editing, value: e.target.value });
          }}
          onKeyDown={onEditKey}
        />
      </div>

      <Grid
        sheet={sheet}
        si={si}
        ev={ev}
        sel={sel}
        onSelect={select}
        editing={editing}
        onEditChange={(value) => editing && setEditing({ ...editing, value })}
        onEditKey={onEditKey}
        onStartEdit={() => startEdit()}
        onFill={onFill}
        onColWidth={(c, w) => act((b) => setColWidth(b, si, c, w))}
        onRowHeight={(r, h) => act((b) => setRowHeight(b, si, r, h))}
        onKeyDown={onKeyDown}
        onBeforeInput={(e) => {
          const data = (e.nativeEvent as InputEvent).data;
          if (editing || !data) return;
          e.preventDefault();
          startEdit(data);
        }}
        onCopy={onCopy}
        onPaste={onPaste}
        rootRef={rootRef}
        zoom={zoom}
        onZoom={(f) => setZoom((z) => Math.min(4, Math.max(0.1, z * f)))}
      />

      <div className="xe-statusbar">
        <SheetTabs
          names={book.sheets.map((x) => x.name)}
          active={si}
          onSelect={(i) => {
            if (i === si) return;
            flush();
            dispatch({ type: 'silent', fn: (b) => ({ ...b, active: i }) });
            setSel({ ar: 0, ac: 0, fr: 0, fc: 0 });
          }}
          onAdd={() => {
            act((b) => addSheet(b));
            setSel({ ar: 0, ac: 0, fr: 0, fc: 0 });
          }}
          onMove={(from, to) => act((b) => moveSheet(b, from, to))}
          onRename={(i, name) => act((b) => renameSheet(b, i, name))}
          onDelete={(i) => {
            act((b) => deleteSheet(b, i));
            setSel({ ar: 0, ac: 0, fr: 0, fc: 0 });
          }}
        />
        <div className="xe-zoom" onMouseDown={(e) => !(e.target instanceof HTMLInputElement) && e.preventDefault()}>
          {stats && (
            <span className="xe-zstats">
              Sum: {+stats.sum.toPrecision(12)} · Average: {+stats.avg.toPrecision(8)} · Count: {stats.n}
            </span>
          )}
          <button title="Zoom out" onClick={() => setZoom((z) => Math.max(0.1, Math.round(z * 20 - 1) / 20))}>−</button>
          <input type="range" min={10} max={400} step={5} value={Math.round(zoom * 100)} onChange={(e) => setZoom(Number(e.target.value) / 100)} aria-label="Zoom" />
          <button title="Zoom in" onClick={() => setZoom((z) => Math.min(4, Math.round(z * 20 + 1) / 20))}>＋</button>
          <button className="xe-pct" title="Reset to 100%" onClick={() => setZoom(1)}>{Math.round(zoom * 100)}%</button>
        </div>
      </div>
    </div>
  );
}
