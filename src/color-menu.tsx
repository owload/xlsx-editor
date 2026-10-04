import { useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { SHADES, STANDARD_COLORS, THEME_COLORS, normalizeHex, shade } from './lib/colors';

export interface ColorMenuProps {
  /** "Fill color" or "Text color": the tooltip, and the name of the swatches for screen readers. */
  title: string;
  /** What the icon shows. */
  icon: string;
  /** The color the button applies before the user has picked another one. */
  initial: string;
  /** The color of the active cell, if it has one: marked in the palette. */
  current: string | undefined;
  /** Colors the workbook already uses, most used first. */
  used: string[];
  /** The label of the item that removes the color ("No fill", "Automatic"). */
  clearLabel: string;
  onPick(color: string): void;
  onClear(): void;
}

const MENU_WIDTH = 222;

function Swatch(props: { hex: string; label: string; title: string; selected: boolean; onPick(hex: string): void }) {
  return (
    <button
      className={'xe-sw' + (props.selected ? ' xe-sel-on' : '')}
      style={{ background: props.hex }}
      title={props.title}
      aria-label={`${props.label} ${props.title}`}
      onClick={() => props.onPick(props.hex)}
    />
  );
}

/**
 * A color button the way spreadsheet programs have it: the button applies the last color picked, the arrow
 * opens a palette of theme colors with shades, standard colors, the colors the workbook already uses, and a
 * custom color (the system picker, or a hex code). No color is stored anywhere but in the workbook.
 */
export function ColorMenu(p: ColorMenuProps) {
  const [open, setOpen] = useState(false);
  const [last, setLast] = useState(p.initial);
  const [hex, setHex] = useState('');
  const [bad, setBad] = useState(false);
  const [pos, setPos] = useState<CSSProperties>({});
  const anchor = useRef<HTMLSpanElement>(null);
  const menu = useRef<HTMLDivElement>(null);
  const native = useRef<HTMLInputElement>(null);

  const close = () => setOpen(false);
  const pick = (color: string) => {
    setLast(color);
    p.onPick(color);
    close();
  };

  // The native picker only reports the color once the user has chosen it ("change"), so dragging in it
  // does not fill the undo history with one step per pixel.
  useEffect(() => {
    const input = native.current;
    if (!open || !input) return;
    const onChange = () => pick(input.value);
    input.addEventListener('change', onChange);
    return () => input.removeEventListener('change', onChange);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  useEffect(() => {
    if (!open) return;
    const outside = (e: MouseEvent) => {
      if (!menu.current?.contains(e.target as Node) && !anchor.current?.contains(e.target as Node)) close();
    };
    const key = (e: KeyboardEvent) => e.key === 'Escape' && close();
    // The menu is placed with fixed coordinates, so it would float away from a ribbon that scrolls or a window
    // that resizes.
    document.addEventListener('mousedown', outside);
    document.addEventListener('keydown', key);
    window.addEventListener('resize', close);
    window.addEventListener('scroll', close, true);
    return () => {
      document.removeEventListener('mousedown', outside);
      document.removeEventListener('keydown', key);
      window.removeEventListener('resize', close);
      window.removeEventListener('scroll', close, true);
    };
  }, [open]);

  const toggle = () => {
    if (!open) {
      const r = anchor.current?.getBoundingClientRect();
      const left = Math.max(4, Math.min(r?.left ?? 0, window.innerWidth - MENU_WIDTH - 4));
      setPos({ left, top: (r?.bottom ?? 0) + 2 });
      setHex('');
      setBad(false);
    }
    setOpen(!open);
  };

  const applyHex = () => {
    const color = normalizeHex(hex);
    if (!color) return setBad(true);
    pick(color);
  };

  const current = p.current?.toLowerCase();
  const swatch = (c: string, label: string, title: string) => (
    <Swatch key={label + c} hex={c} label={p.title} title={title} selected={c === current} onPick={pick} />
  );

  return (
    <span className="xe-colorgrp" ref={anchor}>
      <button className="xe-rb xe-color" title={p.title} aria-label={p.title} onClick={() => pick(last)}>
        <span className="xe-cl">{p.icon}</span>
        <i style={{ background: last }} />
      </button>
      <button className="xe-rb xe-tiny" title={`${p.title}: more colors`} aria-label={`${p.title}: more colors`} aria-expanded={open} aria-haspopup="menu" onClick={toggle}>
        ▾
      </button>
      {open && (
        <div className="xe-cm" role="menu" aria-label={p.title} style={{ ...pos, width: MENU_WIDTH }} ref={menu}>
          <div className="xe-cm-title">Theme colors</div>
          <div className="xe-cm-grid">
            {THEME_COLORS.map((c) => swatch(c.hex, 'theme', c.name))}
            {SHADES.map((s) => THEME_COLORS.map((c) => swatch(shade(c.hex, s.amount), `theme ${s.label}`, `${c.name}, ${s.label.toLowerCase()}`)))}
          </div>
          <div className="xe-cm-title">Standard colors</div>
          <div className="xe-cm-grid">{STANDARD_COLORS.map((c) => swatch(c.hex, 'standard', c.name))}</div>
          {p.used.length > 0 && (
            <>
              <div className="xe-cm-title">Used in this workbook</div>
              <div className="xe-cm-grid">{p.used.map((c) => swatch(c, 'used', c.toUpperCase()))}</div>
            </>
          )}
          <button className="xe-cm-item" onClick={() => { p.onClear(); close(); }}>
            {p.clearLabel}
          </button>
          <label className="xe-cm-item xe-cm-more">
            More colors…
            <input type="color" ref={native} defaultValue={last} aria-label={`${p.title}: choose any color`} />
          </label>
          <div className="xe-cm-hex">
            <input
              type="text"
              value={hex}
              placeholder="#RRGGBB"
              aria-label={`${p.title}: color code`}
              aria-invalid={bad}
              className={bad ? 'xe-bad' : ''}
              maxLength={7}
              onChange={(e) => { setHex(e.target.value); setBad(false); }}
              onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); applyHex(); } }}
            />
            <button onClick={applyHex} disabled={!hex.trim()}>Apply</button>
          </div>
        </div>
      )}
    </span>
  );
}
