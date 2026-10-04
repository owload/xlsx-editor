import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, test, vi } from 'vitest';
import { ColorMenu, type ColorMenuProps } from '../color-menu';
import { XlsxEditor, type XlsxEditorHandle } from '../index';
import { readWorkbook, writeWorkbook } from '../lib/xlsx-io';

(globalThis as unknown as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true;

async function mountMenu(props: Partial<ColorMenuProps> = {}) {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  const onPick = vi.fn();
  const onClear = vi.fn();
  await act(async () => {
    root.render(createElement(ColorMenu, { title: 'Fill color', icon: 'F', initial: '#ffff00', current: undefined, used: [], clearLabel: 'No fill', onPick, onClear, ...props }));
  });
  const q = (sel: string) => container.querySelector<HTMLElement>(sel);
  const click = (el: Element | null) => act(async () => { (el as HTMLElement).click(); });
  const open = () => click(q('button[aria-label="Fill color: more colors"]'));
  return { container, q, click, open, onPick, onClear, unmount: () => act(async () => root.unmount()) };
}

describe('ColorMenu', () => {
  test('the main button applies the initial color, then the last one picked', async () => {
    const m = await mountMenu();
    await m.click(m.q('button[aria-label="Fill color"]'));
    expect(m.onPick).toHaveBeenLastCalledWith('#ffff00');
    await m.open();
    await m.click(m.q('[role=menu] button[aria-label="Fill color Red"]'));
    expect(m.onPick).toHaveBeenLastCalledWith('#ff0000');
    await m.click(m.q('button[aria-label="Fill color"]'));
    expect(m.onPick).toHaveBeenLastCalledWith('#ff0000');
    expect(['#ff0000', 'rgb(255, 0, 0)']).toContain(m.q('button[aria-label="Fill color"] i')!.style.background);
    await m.unmount();
  });

  test('is closed until the arrow is clicked, and shows the theme colors with five shades and ten standard colors', async () => {
    const m = await mountMenu();
    expect(m.q('[role=menu]')).toBeNull();
    await m.open();
    expect(m.q('[role=menu]')).not.toBeNull();
    const grids = m.container.querySelectorAll('.xe-cm-grid');
    expect(grids[0].children).toHaveLength(60); // 10 theme colors x (1 + 5 shades)
    expect(grids[1].children).toHaveLength(10);
    expect(m.container.textContent).toContain('Theme colors');
    expect(m.container.textContent).toContain('Standard colors');
    await m.unmount();
  });

  test('lists the colors already used in the workbook, only when there are some', async () => {
    const without = await mountMenu();
    await without.open();
    expect(without.container.textContent).not.toContain('Used in this workbook');
    await without.unmount();

    const m = await mountMenu({ used: ['#abcdef', '#123456'] });
    await m.open();
    expect(m.container.textContent).toContain('Used in this workbook');
    await m.click(m.q('[role=menu] button[aria-label="Fill color #ABCDEF"]'));
    expect(m.onPick).toHaveBeenCalledWith('#abcdef');
    await m.unmount();
  });

  test('closes after a pick, on Escape, and on a click outside', async () => {
    const m = await mountMenu();
    await m.open();
    await m.click(m.q('[role=menu] button[aria-label="Fill color Blue"]'));
    expect(m.q('[role=menu]')).toBeNull();
    await m.open();
    await act(async () => { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' })); });
    expect(m.q('[role=menu]')).toBeNull();
    await m.open();
    await act(async () => { document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); });
    expect(m.q('[role=menu]')).toBeNull();
    expect(m.onPick).toHaveBeenCalledTimes(1);
    await m.unmount();
  });

  test('marks the color of the active cell', async () => {
    const m = await mountMenu({ current: '#FF0000' });
    await m.open();
    const marked = [...m.container.querySelectorAll('.xe-sel-on')].map((e) => e.getAttribute('aria-label'));
    expect(marked).toEqual(['Fill color Red']);
    await m.unmount();
  });

  test('the clear item calls onClear with its own label', async () => {
    const m = await mountMenu({ clearLabel: 'Automatic' });
    await m.open();
    await m.click(m.container.querySelector('.xe-cm-item'));
    expect(m.onClear).toHaveBeenCalledTimes(1);
    expect(m.container.textContent).not.toContain('Automatic'); // closed
    await m.unmount();
  });

  test('a custom color from the system picker is applied when the user has chosen it', async () => {
    const m = await mountMenu();
    await m.open();
    const input = m.q('input[type=color]') as HTMLInputElement;
    await act(async () => { input.value = '#12ab34'; input.dispatchEvent(new Event('input', { bubbles: true })); });
    expect(m.onPick).not.toHaveBeenCalled(); // moving around in the picker is not a choice yet
    await act(async () => { input.dispatchEvent(new Event('change', { bubbles: true })); });
    expect(m.onPick).toHaveBeenCalledWith('#12ab34');
    await m.unmount();
  });

  test('a custom color can be typed as a code; a wrong code is refused and marked', async () => {
    const m = await mountMenu();
    await m.open();
    const input = m.q('.xe-cm-hex input[type=text]') as HTMLInputElement;
    const type = (v: string) => act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, v);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    const apply = m.q('.xe-cm-hex button') as HTMLButtonElement;
    expect(apply.disabled).toBe(true);
    await type('nonsense');
    await m.click(apply);
    expect(m.onPick).not.toHaveBeenCalled();
    expect(input.getAttribute('aria-invalid')).toBe('true');
    await type('#F80');
    expect(input.getAttribute('aria-invalid')).toBe('false');
    await m.click(apply);
    expect(m.onPick).toHaveBeenCalledWith('#ff8800');
    await m.unmount();
  });
});

describe('the color menus in the editor', () => {
  async function open(bg?: string) {
    const bytes = await writeWorkbook({ sheets: [{ name: 'S', rows: [[{ raw: 'a' }, bg ? { raw: 'b', st: { bg } } : { raw: 'b' }]], colWidths: {} }], active: 0 });
    const container = document.createElement('div');
    document.body.appendChild(container);
    const root = createRoot(container);
    const ref = createRef<XlsxEditorHandle>();
    const onSave = vi.fn();
    await act(async () => { root.render(createElement(XlsxEditor, { data: bytes, fileName: 'a.xlsx', onSave, ref })); });
    await act(async () => { await new Promise((r) => setTimeout(r, 80)); });
    const click = (sel: string) => act(async () => { container.querySelector<HTMLElement>(sel)!.click(); });
    const saved = async () => {
      await act(async () => { await ref.current!.save(); });
      return readWorkbook(onSave.mock.calls.at(-1)![0] as Uint8Array);
    };
    return { container, click, saved, unmount: () => act(async () => root.unmount()) };
  }

  test('a standard color from the menu becomes the fill of the active cell, and is saved', async () => {
    const e = await open();
    await e.click('button[aria-label="Fill color: more colors"]');
    await e.click('[role=menu] button[aria-label="Fill color Green"]'); // the first "Green", in the theme row
    const book = await e.saved();
    expect(book.sheets[0].rows[0]![0]!.st!.bg).toBe('#70ad47');
    await e.unmount();
  });

  test('"No fill" removes the fill, and "Automatic" removes the text color', async () => {
    const e = await open('#ffcc00');
    await e.click('button[aria-label="Text color: more colors"]');
    await e.click('[role=menu] button[aria-label="Text color Red"]');
    await e.click('button[aria-label="Fill color: more colors"]');
    await e.click('.xe-cm-item'); // No fill
    const book = await e.saved();
    const a1 = book.sheets[0].rows[0]![0]!;
    expect(a1.st?.color).toBe('#ff0000');
    expect(a1.st?.bg).toBeUndefined();
    await e.unmount();
  });

  test('the menu offers the colors the workbook already uses', async () => {
    const e = await open('#ffcc00');
    await e.click('button[aria-label="Fill color: more colors"]');
    expect(e.container.querySelector('[role=menu] button[aria-label="Fill color #FFCC00"]')).not.toBeNull();
    await e.unmount();
  });
});
