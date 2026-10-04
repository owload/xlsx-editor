# CLAUDE.md

Guidance for Claude Code when working in this repository.

## What this is

`@owload/xlsx-editor` — a standalone React component for editing `.xlsx` files. It is embedded into **owload**
(an end-to-end encrypted file storage) which decrypts a file, passes the bytes in through `data`, and receives the
edited bytes back through the `onSave` callback. The host repository (`owload-front`) is separate; never edit it
from here. See `README.md` for the public API.

## Rules

- **Do not commit or push unless the owner explicitly asks** (each time).
- **English only** in code, comments, UI strings, docs and commit messages (the host project requires it).
- **Security comes first** — plaintext user data flows through this code:
  - No runtime dependencies besides `react` (peer). Do not add any; prefer a small in-repo implementation. Dev
    dependencies are fine.
  - No network access, no `eval`/`Function`, no `innerHTML`, no storage (`localStorage`, IndexedDB, cookies), no
    logging of cell contents.
  - Treat the input file as hostile: keep the ZIP/XML limits in `src/lib/zip.ts` and `src/lib/xlsx-io.ts`, reject
    `DOCTYPE`/`ENTITY`, validate colors (`#rrggbb`) coming from the file.
  - The only way data leaves the editor is `onSave` (and the system clipboard unless `internalClipboardOnly`).
- Keep the component self-contained: no imports from the host app (`@/...`), no global CSS. Every class is prefixed
  `xe-` and every selector is scoped under `.xe`, so it cannot collide with Tailwind or the host styles.
- Excel-like behaviour and look are the goal (ribbon, grid, tabs, shortcuts). Match Excel when in doubt.

## Commands

```bash
npm run dev        # demo app (demo/) with the editor mounted from src/
npm test           # Vitest + happy-dom
npm run lint
npm run typecheck
npm run build      # dist/: JS, style.css, .d.ts
```

Use Node >= 20 (`.nvmrc` = 22). The default `node` on this machine may be older; `/opt/homebrew/bin/node` is newer.

## As an extension

`src/extension.ts` is the descriptor for `@owload/editor-sdk` (a plain object, type-only import: the SDK is a dev
dependency). `src/lib/inspect.ts` implements `inspect()`; keep its list in step with what `xlsx-io.ts` drops on save
(the "Known limitations" below). The contract and its rules are in the SDK's README; `src/test/conformance.test.tsx`
runs the SDK's suite and must keep passing. `src/lib/preview.ts` draws the PNG preview (`layoutSheetPreview` is pure and tested;
`renderSheetPreview` draws through the adapter in `src/lib/canvas.ts`). The component has a close button in its title bar (`aria-label="Close"`, shown when `onClose` is given); it only calls `onClose` — the host asks about unsaved changes and removes the editor.

## Architecture

- `src/lib/` is framework-free: `xlsx-io.ts` (reader/writer; `DOMParser` + own `zip.ts` on
  `CompressionStream`/`DecompressionStream`), `formula.ts` (tokenizer, parser, evaluator, reference rewriting),
  `numfmt.ts` (Excel number/date formats), `ops.ts` (styles, fill, sort, find/replace, autosum), `store.ts`
  (immutable book operations + undo/redo reducer with dirty tracking), `types.ts`.
- The workbook model is `Book -> Sheet -> rows[r][c] = Cell { raw, z?, text?, st? }`. `raw` is exactly what the
  formula bar shows (`=...` for formulas). Operations return new objects (structural sharing) so undo is cheap.
- `xlsx-editor.tsx` owns state (history reducer, selection, editing, zoom), keyboard and clipboard, and the
  load/save lifecycle (`data` is read once on mount; `onSave` receives serialized bytes; dirty tracking via
  `rev`/`savedRev`). `grid.tsx` renders a virtualized grid with absolutely positioned cells and handles mouse
  selection, header selection, fill handle and Ctrl+wheel zoom.
- React hooks lint note: keep render loops out of components as `for`/`while` statements (use
  `Array.from(...).map/flatMap`), otherwise `react-hooks/rules-of-hooks` reports false "called in a loop" errors.

## Known limitations (data dropped on save)

Charts, images, conditional formatting, data validation, defined names, comments, freeze panes, row heights,
theme/indexed colors, array formulas. Keep `README.md` in sync when this changes.
