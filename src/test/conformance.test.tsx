import { act, runConformanceTests } from '@owload/editor-sdk/testing';
import { extension } from '../extension';
import { writeWorkbook } from '../lib/xlsx-io';

const sample = await writeWorkbook({
  sheets: [{ name: 'Sheet1', rows: [[{ raw: 'start' }]], colWidths: {} }],
  active: 0,
});

let counter = 0;
const bar = (container: HTMLElement) => container.querySelector<HTMLInputElement>('.xe-formulabar input')!;

// Edits go through the formula bar: it commits to the active cell with Enter, which then moves down,
// so every call changes a different cell.
runConformanceTests(extension, {
  sample,
  ready: (container) => !!container.querySelector('.xe-grid'),
  focusTarget: (container) => bar(container),
  // One gesture per act(): the editor has to render the edit state before the next one arrives.
  async edit(container) {
    const input = bar(container);
    await act(async () => { input.focus(); });
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, `edited${++counter}`);
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    });
  },
  verifyReopened(container) {
    // The first edit went to A1 (the active cell on open).
    const value = bar(container).value;
    if (!value.startsWith('edited')) throw new Error(`A1 holds "${value}" after reopening, not the edit.`);
  },
});
