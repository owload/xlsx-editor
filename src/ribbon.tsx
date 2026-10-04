import { ColorMenu } from './color-menu';
import type { Style } from './lib/types';

export interface RibbonActions {
  save(): void;
  undo(): void;
  redo(): void;
  patch(p: Partial<Style>): void;
  setFormat(z: string | undefined): void;
  decimals(d: 1 | -1): void;
  autosum(): void;
  sort(asc: boolean): void;
  find(replace: boolean): void;
  insertRows(below: boolean): void;
  deleteRows(): void;
  insertCols(right: boolean): void;
  deleteCols(): void;
  clear(): void;
}

interface Props {
  st: Style | undefined;
  z: string | undefined;
  canUndo: boolean;
  canRedo: boolean;
  fileName: string;
  dirty: boolean;
  saving: boolean;
  /** Colors the workbook already uses, for the color menus. */
  usedColors: string[];
  onClose?: () => void;
  a: RibbonActions;
}

const FORMATS: [string, string | undefined][] = [
  ['General', undefined],
  ['Number', '0.00'],
  ['Currency', '"$"#,##0.00'],
  ['Percentage', '0%'],
  ['Date', 'yyyy-mm-dd'],
  ['Time', 'h:mm:ss'],
  ['Text', '@'],
];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

function formatName(z: string | undefined) {
  if (!z) return 'General';
  const hit = FORMATS.find(([, f]) => f === z);
  return hit ? hit[0] : 'custom';
}

function Btn(props: { title: string; on?: boolean; disabled?: boolean; onClick: () => void; children: React.ReactNode; wide?: boolean }) {
  return (
    <button
      className={'xe-rb' + (props.on ? ' xe-on' : '') + (props.wide ? ' xe-wide' : '')}
      title={props.title}
      disabled={props.disabled}
      onClick={props.onClick}
    >
      {props.children}
    </button>
  );
}

export function Ribbon({ st, z, canUndo, canRedo, fileName, dirty, saving, usedColors, onClose, a }: Props) {
  const s = st ?? {};
  return (
    <div className="xe-ribbon-wrap" onMouseDown={(e) => !(e.target instanceof HTMLSelectElement) && !(e.target instanceof HTMLInputElement) && e.preventDefault()}>
      <div className="xe-titlebar">
        <div className="xe-qat">
          <button className="xe-rb xe-dark" title="Save (Ctrl+S)" disabled={saving} onClick={a.save}>💾</button>
          <button className="xe-rb xe-dark" title="Undo (Ctrl+Z)" disabled={!canUndo} onClick={a.undo}>↶</button>
          <button className="xe-rb xe-dark" title="Redo (Ctrl+Y)" disabled={!canRedo} onClick={a.redo}>↷</button>
        </div>
        <div className="xe-doc">
          {fileName}
          {dirty ? ' — unsaved changes' : ''}
        </div>
        <button className="xe-savebtn" onClick={a.save} disabled={saving}>
          {saving ? 'Saving…' : 'Save'}
        </button>
        {onClose && (
          <button className="xe-rb xe-dark xe-close" aria-label="Close" title="Close" disabled={saving} onClick={onClose}>
            ✕
          </button>
        )}
      </div>
      <div className="xe-tabstrip">
        <span className="xe-rtab xe-cur">Home</span>
      </div>
      <div className="xe-ribbon">
        <div className="xe-group">
          <div className="xe-row">
            <select className="xe-sel" value={s.sz ?? 11} title="Font size" onChange={(e) => a.patch({ sz: Number(e.target.value) === 11 ? undefined : Number(e.target.value) })}>
              {SIZES.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </div>
          <div className="xe-row">
            <Btn title="Bold (Ctrl+B)" on={!!s.b} onClick={() => a.patch({ b: !s.b })}><b>B</b></Btn>
            <Btn title="Italic (Ctrl+I)" on={!!s.i} onClick={() => a.patch({ i: !s.i })}><i>I</i></Btn>
            <Btn title="Underline (Ctrl+U)" on={!!s.u} onClick={() => a.patch({ u: !s.u })}><u>U</u></Btn>
            <Btn title="All borders" on={!!s.bd} onClick={() => a.patch({ bd: !s.bd })}>▦</Btn>
            <ColorMenu title="Fill color" icon="🪣" initial="#ffff00" current={s.bg} used={usedColors} clearLabel="No fill" onPick={(c) => a.patch({ bg: c })} onClear={() => a.patch({ bg: undefined })} />
            <ColorMenu title="Text color" icon="A" initial="#ff0000" current={s.color} used={usedColors} clearLabel="Automatic" onPick={(c) => a.patch({ color: c })} onClear={() => a.patch({ color: undefined })} />
          </div>
          <div className="xe-glabel">Font</div>
        </div>
        <div className="xe-group">
          <div className="xe-row">
            <Btn title="Align left" on={s.h === 'left'} onClick={() => a.patch({ h: s.h === 'left' ? undefined : 'left' })}>⇤</Btn>
            <Btn title="Center" on={s.h === 'center'} onClick={() => a.patch({ h: s.h === 'center' ? undefined : 'center' })}>↔</Btn>
            <Btn title="Align right" on={s.h === 'right'} onClick={() => a.patch({ h: s.h === 'right' ? undefined : 'right' })}>⇥</Btn>
          </div>
          <div className="xe-glabel">Alignment</div>
        </div>
        <div className="xe-group">
          <div className="xe-row">
            <select className="xe-sel xe-wide" value={formatName(z)} title="Number format" onChange={(e) => a.setFormat(FORMATS.find(([n]) => n === e.target.value)?.[1])}>
              {formatName(z) === 'custom' && <option value="custom">{z}</option>}
              {FORMATS.map(([n]) => (
                <option key={n}>{n}</option>
              ))}
            </select>
          </div>
          <div className="xe-row">
            <Btn title="Currency format" onClick={() => a.setFormat('"$"#,##0.00')}>$</Btn>
            <Btn title="Percent format" onClick={() => a.setFormat('0%')}>%</Btn>
            <Btn title="Thousands separator" onClick={() => a.setFormat('#,##0.00')}>000</Btn>
            <Btn title="Increase decimals" onClick={() => a.decimals(1)}>.0→</Btn>
            <Btn title="Decrease decimals" onClick={() => a.decimals(-1)}>←.0</Btn>
          </div>
          <div className="xe-glabel">Number</div>
        </div>
        <div className="xe-group">
          <div className="xe-row">
            <Btn wide title="Insert row above" onClick={() => a.insertRows(false)}>＋ Row above</Btn>
            <Btn wide title="Insert column left" onClick={() => a.insertCols(false)}>＋ Column left</Btn>
          </div>
          <div className="xe-row">
            <Btn wide title="Insert row below" onClick={() => a.insertRows(true)}>＋ Row below</Btn>
            <Btn wide title="Insert column right" onClick={() => a.insertCols(true)}>＋ Column right</Btn>
          </div>
          <div className="xe-row">
            <Btn wide title="Delete rows" onClick={a.deleteRows}>✕ Rows</Btn>
            <Btn wide title="Delete columns" onClick={a.deleteCols}>✕ Columns</Btn>
          </div>
          <div className="xe-glabel">Cells</div>
        </div>
        <div className="xe-group">
          <div className="xe-row">
            <Btn wide title="AutoSum" onClick={a.autosum}>Σ AutoSum</Btn>
            <Btn wide title="Clear contents (Delete)" onClick={a.clear}>⌫ Clear</Btn>
          </div>
          <div className="xe-row">
            <Btn wide title="Sort A to Z" onClick={() => a.sort(true)}>A→Z</Btn>
            <Btn wide title="Sort Z to A" onClick={() => a.sort(false)}>Z→A</Btn>
          </div>
          <div className="xe-row">
            <Btn wide title="Find (Ctrl+F)" onClick={() => a.find(false)}>🔍 Find</Btn>
            <Btn wide title="Replace (Ctrl+H)" onClick={() => a.find(true)}>Replace</Btn>
          </div>
          <div className="xe-glabel">Editing</div>
        </div>
      </div>
    </div>
  );
}
