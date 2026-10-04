import type { EditorExtension } from '@owload/editor-sdk';

/**
 * The descriptor the host loads eagerly (owload-docs/decisions/0019). A plain object: only the type
 * comes from the SDK, so the package has no runtime dependency; the host validates it. The editor and
 * `inspect` are separate chunks, loaded when a spreadsheet is opened.
 *
 * `readOnly` is not supported yet (the contract's property is optional), so the host opens
 * spreadsheets for editing only.
 */
export const extension = {
  apiVersion: 1,
  id: 'xlsx',
  label: 'Spreadsheet',
  fileExtensions: ['xlsx'],
  createNew: { label: 'spreadsheet', defaultExtension: 'xlsx' },
  maxFileBytes: 100 * 1024 * 1024,
  load: () => import('./xlsx-editor').then((m) => ({ default: m.XlsxEditor })),
  inspect: (data: Uint8Array) => import('./lib/inspect').then((m) => m.inspectWorkbook(data)),
  // The top-left corner of the active sheet, drawn lazily; null where there is no canvas or no data.
  preview: (data: Uint8Array, options: { size: number }) =>
    import('./lib/preview').then((m) => m.renderSheetPreview(data, options.size)),
} satisfies EditorExtension;
