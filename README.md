# @owload/xlsx-editor

A self-contained spreadsheet (`.xlsx`) editor component. It has **no dependencies besides React** and
makes **no network requests**: the host hands it the file bytes and receives the edited bytes back through
a callback. Nothing is stored, logged or sent anywhere by the module itself.

## Install

Until it is published to a registry, depend on a pinned release tag straight from GitHub (the `prepare`
script builds `dist/` on install):

```json
"@owload/xlsx-editor": "github:<owner>/xlsx-editor#v0.1.0"
```

Pin an exact tag (never a branch) and review the diff on every bump. The host must provide React 19
(`peerDependencies`); with Vite add `resolve.dedupe: ['react', 'react-dom']` to avoid two React copies.

## As an Owload extension

The package follows the [`@owload/editor-sdk`](https://github.com/owload/editor-sdk) contract
([ADR 0019](https://github.com/owload/owload-docs/blob/main/decisions/0019-editor-extensions.md)), so the Owload client
plugs it in through a descriptor instead of rendering the component by hand:

```ts
import { extension } from "@owload/xlsx-editor/extension"; // id "xlsx", ".xlsx", "New spreadsheet", 100 MiB
import "@owload/xlsx-editor/style.css";                      // imported lazily by the host, with the editor
```

- `extension.load()` returns the editor as a separate chunk; the descriptor itself is a few lines.
- `extension.inspect(bytes)` reads a file and lists what saving it with this editor would drop (charts, images,
  comments, conditional formatting, frozen panes, row heights, …) as `{ id, label }` entries, or an empty list
  when nothing would be lost. The host shows it before the user edits.
- `readOnly` is not supported yet; the property is optional in the contract.
- Ctrl/Cmd+S saves from everywhere in the editor, also while a cell or the formula bar is being edited.
- The conformance suite of the SDK runs in `src/test/conformance.test.tsx`.

## Usage

```tsx
import { XlsxEditor } from "@owload/xlsx-editor";
import "@owload/xlsx-editor/style.css"; // once, anywhere in the app

<XlsxEditor
  key={file.id}                 // remount to open another file
  data={plaintextBytes}         // Uint8Array | ArrayBuffer | null (null/empty = blank workbook)
  fileName={file.name}          // display only
  onSave={async (bytes) => {    // called on Save / Ctrl+S with a complete .xlsx
    await saveFile(new File([bytes], file.name), pwd()!);
  }}
  onDirtyChange={setDirty}      // optional: drive your own "unsaved changes" dialog
  onError={(e) => console.error(e)} // optional: load/save failures (also shown in the UI)
  internalClipboardOnly         // optional: never touch the system clipboard
/>
```

The component fills its parent (`height: 100%`), so give the parent a definite height, e.g. a full-screen
modal container.

### Props

| Prop | Description |
| --- | --- |
| `data` | File contents. Read **once on mount** and never kept afterwards. To open another file, change the component `key`. Corrupt or oversized input shows an error panel instead of the editor, so a broken file is never replaced by a blank workbook. |
| `fileName` | Shown in the title bar only. The module never reads or writes files by name. |
| `onSave(bytes)` | Receives the serialized `.xlsx`. While it is pending the Save button shows "Saving…"; if it throws/rejects the error is shown and the document stays dirty. Edits made while saving keep the document dirty. |
| `onDirtyChange(dirty)` | Fires when the unsaved-changes state flips. |
| `onError(error)` | Fires on load/save failure. |
| `maxFileBytes` | Reject larger input. Default 100 MiB. |
| `internalClipboardOnly` | Copy/cut write an empty string to the system clipboard and paste ignores it; cells are kept in memory inside the editor (formulas and formatting included). Default `false`. |
| `className` | Extra class on the root element. |
| `ref` | `XlsxEditorHandle`: `save(): Promise<void>` and `isDirty(): boolean`, e.g. to save before closing. |

The module does not warn on page unload and does not show a "discard changes?" dialog; use
`onDirtyChange` / `ref.isDirty()` for that, the way a plain text editor would.

## What is supported

- Reading/writing `.xlsx`: values, formulas (with cached results), number formats, fonts (bold, italic,
  underline, size, color), fills, thin borders, horizontal alignment, column widths, merged cells, multiple sheets.
- Editing: formula bar, undo/redo (100 steps), copy/cut/paste (also TSV from Excel), fill handle, find/replace,
  sort, AutoSum, insert/delete rows and columns (formulas are rewritten), multi-range selection, zoom,
  sheet tabs (add, rename, delete, drag to reorder).
- Formulas: `+ - * / ^ & %`, comparisons, cross-sheet references, and `SUM AVERAGE MIN MAX COUNT COUNTA
  PRODUCT IF IFERROR AND OR NOT ROUND ABS SQRT INT MOD POWER LEN UPPER LOWER TRIM LEFT RIGHT CONCAT(ENATE)`.
  Unknown functions evaluate to `#NAME?`; formulas are parsed by a small hand-written parser and are never
  executed as code.

## Known limitations

Data the module does not understand is **dropped on save** (the file is rebuilt from the parsed model):
charts, images, conditional formatting, data validation, defined names, comments, freeze panes, row heights,
theme/indexed colors, fonts other than the default, array formulas, pivot tables, macros. Editing a file that
relies on these will lose them; consider warning the user or offering a copy.

## Security notes

- Parsing is bounded: ≤ 2000 archive entries, ≤ 200 MiB per entry and ≤ 500 MiB total, decompression is
  capped at the declared size, CRC-32 is verified, encrypted ZIPs are rejected.
- XML with a `DOCTYPE`/`ENTITY` is rejected. Colors read from the file must match `#rrggbb`.
- All output is escaped by React (no `innerHTML`, no `eval`, no `Function`).
- The system clipboard is the one place plaintext can leave the editor; use `internalClipboardOnly` to close it.
- Column-resize/row-header cursors are SVG `data:` URIs in the stylesheet; a strict CSP needs `img-src data:`
  for them (otherwise the browser falls back to the default resize cursors).

## Styling

All CSS is shipped as `dist/style.css` (source: `src/xlsx-editor.css`). Every selector is scoped to `.xe` / `.xe-*`,
so it neither leaks into nor is hit by utility frameworks such as Tailwind. Colors are CSS variables on `.xe`; a
`.dark` ancestor switches to the dark palette.

## Development

Requires Node >= 20 (`.nvmrc` pins 22).

```bash
npm ci
npm run dev        # demo app (demo/): open an .xlsx, edit, save -> bytes are handed to onSave
npm test           # Vitest (happy-dom provides DOMParser)
npm run lint
npm run typecheck
npm run build      # dist/: ES module, style.css, type declarations
```

Layout:

```
src/index.ts          public API (XlsxEditor, XlsxEditorProps, XlsxEditorHandle)
src/xlsx-editor.tsx   main component (state, keyboard, clipboard, save/load lifecycle)
src/grid.tsx          virtualized grid, selection, zoom, fill handle
src/ribbon.tsx        toolbar        src/find-bar.tsx  find/replace      src/sheet-tabs.tsx  sheet tabs
src/lib/              framework-free logic: xlsx-io (reader/writer), zip, formula, numfmt, ops, store, types
src/test/             unit tests
demo/                 demo app for manual testing
```

## Releasing

Before tagging a release run `npm run lint && npm run typecheck && npm test && npm run build`, then tag
`vX.Y.Z` and push the tag. Hosts should depend on the exact tag.
