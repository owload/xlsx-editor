import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, test, vi } from 'vitest';
import { XlsxEditor, type XlsxEditorHandle } from '../index';
import { writeWorkbook } from '../lib/xlsx-io';
import { unzip, zip } from '../lib/zip';

(globalThis as unknown as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true;

const plain = () => writeWorkbook({ sheets: [{ name: 'S', rows: [[{ raw: 'a' }]], colWidths: {} }], active: 0 });

/** A workbook file to which parts the editor cannot keep were added. */
async function withExtras(extra: Record<string, string>) {
  const files = await unzip(await plain());
  for (const [name, content] of Object.entries(extra)) files.set(name, new TextEncoder().encode(content));
  return zip(files);
}

async function open(data: Uint8Array | null) {
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  const ref = createRef<XlsxEditorHandle>();
  const onSave = vi.fn();
  await act(async () => {
    root.render(createElement(XlsxEditor, { data, fileName: 'a.xlsx', onSave, ref }));
  });
  await act(async () => { await new Promise((r) => setTimeout(r, 80)); });
  const note = () => container.querySelector('[role=note]');
  return { container, ref, onSave, note, unmount: () => act(async () => root.unmount()) };
}

describe('the note about what saving would drop', () => {
  test('names what the file holds that the editor cannot keep', async () => {
    const t = await open(await withExtras({ 'xl/charts/chart1.xml': '<c/>', 'xl/media/image1.png': 'x' }));
    expect(t.note()!.textContent).toBe('Saving this file here will drop: charts, images.Got it');
    await t.unmount();
  });

  test('is not shown for a file that loses nothing, or for a new document', async () => {
    const t = await open(await plain());
    expect(t.note()).toBeNull();
    await t.unmount();
    const blank = await open(null);
    expect(blank.note()).toBeNull();
    await blank.unmount();
  });

  test('goes away when dismissed', async () => {
    const t = await open(await withExtras({ 'xl/comments1.xml': '<c/>' }));
    expect(t.note()).not.toBeNull();
    await act(async () => { t.container.querySelector<HTMLButtonElement>('[role=note] button')!.click(); });
    expect(t.note()).toBeNull();
    await t.unmount();
  });

  test('goes away once the file has been saved, because it has been rewritten', async () => {
    const t = await open(await withExtras({ 'xl/comments1.xml': '<c/>' }));
    expect(t.note()).not.toBeNull();
    await act(async () => { await t.ref.current!.save(); });
    expect(t.onSave).toHaveBeenCalledTimes(1);
    expect(t.note()).toBeNull();
    await t.unmount();
  });
});
