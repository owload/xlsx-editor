import { useEffect, useRef, useState } from 'react';

interface Props {
  names: string[];
  active: number;
  onSelect(i: number): void;
  onAdd(): void;
  onRename(i: number, name: string): void;
  onDelete(i: number): void;
  onMove(from: number, to: number): void;
}

export function SheetTabs({ names, active, onSelect, onAdd, onRename, onDelete, onMove }: Props) {
  const strip = useRef<HTMLDivElement>(null);
  const [edge, setEdge] = useState({ left: false, right: false });
  const [renaming, setRenaming] = useState<{ i: number; value: string } | null>(null);
  const [drag, setDrag] = useState<{ from: number; slot: number } | null>(null);
  const justDragged = useRef(false);
  const [menu, setMenu] = useState<{ i: number; x: number; y: number } | null>(null);

  const measure = () => {
    const el = strip.current;
    if (el) setEdge({ left: el.scrollLeft > 0, right: el.scrollLeft + el.clientWidth < el.scrollWidth - 2 });
  };

  useEffect(() => {
    const el = strip.current;
    if (!el) return;
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [names.length]);

  // the active tab and its left/right neighbours must stay in view
  useEffect(() => {
    const el = strip.current;
    if (!el) return;
    const tabs = Array.from(el.querySelectorAll<HTMLElement>('[data-i]'));
    const tab = tabs[active];
    if (!tab) return;
    const prev = tabs[active - 1];
    const next = tabs[active + 1];
    const wantLeft = (prev ?? tab).offsetLeft;
    const wantRight = next ? next.offsetLeft + next.offsetWidth : tab.offsetLeft + tab.offsetWidth;
    if (wantRight > el.scrollLeft + el.clientWidth) el.scrollLeft = wantRight - el.clientWidth;
    if (wantLeft < el.scrollLeft) el.scrollLeft = wantLeft;
    // if the strip is too narrow for the neighbours, at least keep the tab itself in view
    if (tab.offsetLeft < el.scrollLeft) el.scrollLeft = tab.offsetLeft;
    else if (tab.offsetLeft + tab.offsetWidth > el.scrollLeft + el.clientWidth) el.scrollLeft = tab.offsetLeft + tab.offsetWidth - el.clientWidth;
    measure();
  }, [active, names]);

  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    const key = (e: KeyboardEvent) => e.key === 'Escape' && close();
    window.addEventListener('mousedown', close);
    window.addEventListener('keydown', key);
    window.addEventListener('blur', close);
    return () => {
      window.removeEventListener('mousedown', close);
      window.removeEventListener('keydown', key);
      window.removeEventListener('blur', close);
    };
  }, [menu]);

  const scrollBy = (dx: number) => {
    strip.current?.scrollBy({ left: dx });
    measure();
  };
  const scrollEnd = (right: boolean) => {
    const el = strip.current;
    if (el) el.scrollTo({ left: right ? el.scrollWidth : 0 });
    measure();
  };

  /** Tab dragging: the slot is the insertion position 0..n, decided by neighbouring tab midpoints. */
  const startDrag = (i: number, e: React.MouseEvent) => {
    if (e.button !== 0) return;
    const x0 = e.clientX;
    let active = false;
    const slotAt = (x: number) => {
      const tabs = Array.from(strip.current?.querySelectorAll<HTMLElement>('[data-i]') ?? []);
      let slot = 0;
      for (const t of tabs) {
        const r = t.getBoundingClientRect();
        if (x > r.left + r.width / 2) slot++;
      }
      return slot;
    };
    const move = (ev: MouseEvent) => {
      if (!active && Math.abs(ev.clientX - x0) < 5) return;
      active = true;
      const el = strip.current;
      if (el) {
        const r = el.getBoundingClientRect();
        if (ev.clientX < r.left + 30) el.scrollLeft -= 14;
        else if (ev.clientX > r.right - 30) el.scrollLeft += 14;
        measure();
      }
      setDrag({ from: i, slot: slotAt(ev.clientX) });
    };
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (!active) return;
      justDragged.current = true;
      setTimeout(() => (justDragged.current = false), 0);
      const slot = slotAt(ev.clientX);
      const to = slot > i ? slot - 1 : slot;
      if (to !== i) onMove(i, to);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const indicatorLeft = (() => {
    if (!drag) return 0;
    const tabs = Array.from(strip.current?.querySelectorAll<HTMLElement>('[data-i]') ?? []);
    const t = tabs[Math.min(drag.slot, tabs.length - 1)];
    if (!t) return 0;
    return drag.slot >= tabs.length ? t.offsetLeft + t.offsetWidth : t.offsetLeft;
  })();

  const finish = () => {
    if (!renaming) return;
    const { i, value } = renaming;
    setRenaming(null);
    if (value.trim()) onRename(i, value);
  };

  return (
    <div className="xe-sheettabs" onMouseDown={(e) => !(e.target instanceof HTMLInputElement) && e.preventDefault()}>
      <div className="xe-navs">
        <button disabled={!edge.left} title="First sheet" onClick={() => scrollEnd(false)}>⏮</button>
        <button disabled={!edge.left} title="Scroll left" onClick={() => scrollBy(-140)}>◀</button>
        <button disabled={!edge.right} title="Scroll right" onClick={() => scrollBy(140)}>▶</button>
        <button disabled={!edge.right} title="Last sheet" onClick={() => scrollEnd(true)}>⏭</button>
      </div>
      <div
        className="xe-strip"
        ref={strip}
        onScroll={measure}
        onWheel={(e) => {
          if (strip.current && e.deltaY) strip.current.scrollLeft += e.deltaY;
        }}
      >
        {names.map((n, i) =>
          renaming?.i === i ? (
            <input
              key={i}
              data-i={i}
              className="xe-tab-rename"
              autoFocus
              maxLength={31}
              value={renaming.value}
              onFocus={(e) => e.currentTarget.select()}
              onChange={(e) => setRenaming({ i, value: e.target.value })}
              onBlur={finish}
              onKeyDown={(e) => {
                if (e.key === 'Enter') finish();
                if (e.key === 'Escape') setRenaming(null);
              }}
            />
          ) : (
            <button
              key={i}
              data-i={i}
              className={'xe-stab' + (i === active ? ' xe-cur' : '') + (drag?.from === i ? ' xe-dragging' : '')}
              title={n}
              onMouseDown={(e) => startDrag(i, e)}
              onClick={() => !justDragged.current && onSelect(i)}
              onDoubleClick={() => setRenaming({ i, value: n })}
              onContextMenu={(e) => {
                e.preventDefault();
                onSelect(i);
                setMenu({ i, x: e.clientX, y: e.clientY });
              }}
            >
              {n}
            </button>
          ),
        )}
        {drag && <div className="xe-dropmark" style={{ left: indicatorLeft }} />}
      </div>
      <button className="xe-addtab" title="New sheet" onClick={onAdd}>＋</button>
      {menu && (
        <div className="xe-ctxmenu" style={{ left: menu.x, bottom: window.innerHeight - menu.y }} onMouseDown={(e) => e.stopPropagation()}>
          <button onClick={() => { setMenu(null); onAdd(); }}>Insert</button>
          <button onClick={() => { setMenu(null); setRenaming({ i: menu.i, value: names[menu.i] }); }}>Rename</button>
          <button disabled={names.length < 2} onClick={() => { setMenu(null); onDelete(menu.i); }}>Delete</button>
        </div>
      )}
    </div>
  );
}
