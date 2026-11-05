# XLSX Editor

An Orca custom editor for Excel workbooks that draws the formatting the file actually carries —
fills, fonts, borders, alignment, number formats, merged cells, frozen panes and every sheet — and
saves by rewriting only the cells that changed.

Orca hands a custom editor the workbook's bytes and keeps the file's path to itself. The editor
parses those bytes, shows them, and posts back new bytes; Orca owns the buffer, the dirty state and
the save.

## Why the save is byte-preserving

A spreadsheet is a zip of XML parts: styles, shared strings, drawings, charts, macros, pivot
caches. Re-serialising the whole workbook would rewrite all of them, and anything the writer does
not model — `vbaProject.bin`, a chart, a comment — would be lost or invalidated.

So the writer never re-serialises. It edits the `<c>` elements of the edited sheets in the
original XML text, leaves `sharedStrings.xml` alone by writing new text inline, and copies every
other zip entry's **compressed bytes** verbatim. On `test/fixtures/styled.xlsx` — nine parts in
the archive — one cell edit rewrites one of them, `xl/worksheets/sheet1.xml`, and the other eight
stay identical byte for byte.

Recalculation bookkeeping is the exception, and it is conditional:

- when the package carries `xl/calcChain.xml`, the save deletes it and rewrites
  `xl/_rels/workbook.xml.rels` and `[Content_Types].xml` to drop the relationship and the
  content-type override that pointed at it;
- when `<calcPr>` does not already say `fullCalcOnLoad="1"`, the save rewrites `xl/workbook.xml`
  to add it.

So the case `test/writer.spec.mjs` asserts — a calc chain plus a `<calcPr>` without
`fullCalcOnLoad` — comes back with four changed parts plus one removal; a workbook with neither
comes back with two changed parts, the sheet and `xl/workbook.xml`.

## Layout

| Path | What it holds |
| --- | --- |
| `orca-plugin.json` | manifest: `sandboxed-ui`, `contentKind: "bytes"`, `frame: "isolated"` |
| `editor/index.html` | the frame page; the build fills its `vendor-script`, `app-script` and `app-styles` blocks |
| `editor/boot.js` | the classic-script watchdog that reports missing assets instead of spinning on the loader |
| `editor/page.css` | the base shell and theme fallbacks, hand-written and linked by `index.html` |
| `editor/assets/exceljs-4.4.0.min.js` | vendored reader, the `setImmediate` polyfill replaced and five lodash global probes patched |
| `editor/assets/app-<hash>.js`, `.css` | this plugin, bundled |
| `src/main.js` | boot: session, model, chrome |
| `src/ui/document.js` | the only file that opens an Orca document session: decode, encode, save marks |
| `src/ui/app.js` | sheet tabs, formula bar, status bar, undo and redo, the copy and export menus |
| `src/ui/grid.js` | virtualised grid: styles, merges, frozen panes, selection, navigation |
| `src/ui/clipboard.js` | TSV, CSV and Markdown formatters; `copyText` is its only DOM call |
| `src/model/workbook.js` | ExcelJS to a view model, existing cells only |
| `src/model/style.js`, `numfmt.js` | styles to CSS, number formats to displayed text |
| `src/write/xml.js`, `xlsx-writer.js` | text-level XML surgery, cell edits, recalculation bookkeeping |
| `src/write/zip.js`, `zip-budget.js` | zip reading and rewriting that keeps untouched entries' bytes |
| `src/edit/edits.js` | the operation log, coalescing, undo and redo, save marks |
| `scripts/make-fixtures.py` | the fixture generator described below |
| `docs/CONTRACT.md` | the interface between those layers |

## Build and test

```sh
npm ci --ignore-scripts
npm run build        # editor/assets/* and the three build-owned blocks in index.html; fails on eval, on a missing asset, on a file over 1 MiB
npm run check:dist   # the same checks against what is on disk, without bundling or writing
npm test             # node --test over both patterns: 127 tests
orca plugin test .   # what Orca runs: 12 tests, from test/edits.test.mjs
```

`editor/assets/` is committed on purpose: the built bundle, its stylesheet and the vendored ExcelJS
are what an installed plugin loads, so the folder opens nothing without them. `check:dist` is what
catches that drift before a commit.

Two runners, on purpose. `orca plugin test` selects files by name — its help lists `*.test.ts(x)`
and also `.mts`, `.js`, `.jsx`, `.mjs` — so a suite in `*.spec.mjs` is simply invisible to it:
`test/edits.test.mjs` is the only file that matches, and it is dependency-free. Everything that
needs `exceljs` or `fflate` sits in `*.spec.mjs` and runs under plain `node --test`, which globs
both patterns.

`scripts/make-fixtures.py` builds the fixtures with openpyxl, an implementation independent of the
reader, so the tests never check ExcelJS against itself. It writes three files into
`test/fixtures/`: `rich.xlsx` (three sheets, styles, number formats, merged cells, a conditional
format, a frozen pane, a defined name), `macro.xlsm` (the same book plus a fake
`xl/vbaProject.bin`, so the save can be proved to leave the macro part byte for byte) and
`styled.xlsx` (166 rows by 16 columns, a painted header, banded rows, borders everywhere, mixed
number formats). `styled.xlsx` is generated precisely because it stands in for a real customer
workbook, which must not be committed to a public repository.

## Run it in Orca

```sh
orca plugin dev load . --json                        # prints the approval fingerprint; exits 3 while approval is owed
orca plugin dev load . --json --approve <fingerprint>  # same folder, approved: exits 0
orca plugin open local.xlsx-editor --editor ./book.xlsx
orca plugin shot local.xlsx-editor --editor ./book.xlsx --out ./shot.png --json
```

`plugin dev load` keeps the folder in Settings → Plugins → Development folders and reloads it when
its files change. Exit code 3 is `approval_required` (or a fingerprint that the files, version or
access changed out from under, `approval_stale`); once the fingerprint is approved the same command
exits 0. `plugin shot` renders the editor exactly as Orca draws it and writes a PNG.

## What it does not do

- No charts, pivot tables, sparklines or drawing objects: they survive the save untouched, but are
  not drawn.
- No formula evaluation. A formula shows its cached result, the formula text when the file carries
  none; an edited formula is written without a result and Excel recalculates on open.
- No insert or delete of rows and columns yet: that shifts every reference in the file (formulas,
  merges, validations, conditional formats, defined names) and is a feature of its own.
- No `.xls`, `.xlsb` or `.ods`: only the Office Open XML formats this writer can patch.
- Six workbooks out of 478 sampled from a real machine still fail to open: ExcelJS gives up on a
  drawing part with an unprefixed `<wsDr>` root (4 files), one table part (1) and one spread with
  prefixed namespaces throughout (1). Everything else in that corpus opens.
- No network: the frame refuses every request outside the plugin.
