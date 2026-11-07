# Module contract

The interfaces between the slices of this plugin, as the code implements them.

Runtime: browser editor frame (`frame: "isolated"`), bundled by esbuild to `editor/assets/`.
No network, no `eval` / `new Function`, no external requests. Everything is relative to the plugin.

## Slices and the files they own

| Path | Owner |
| --- | --- |
| `build/build.mjs`, `package.json`, `editor/index.html`, `editor/boot.js`, `editor/page.css` | slice A |
| `src/write/zip.js`, `src/write/zip-budget.js` | slice B |
| `src/write/xml.js`, `src/write/xlsx-writer.js`, `src/write/structural.js`, `src/write/structural-formula.js`, `src/write/structural-objects.js` | slice C |
| `src/model/workbook.js`, `src/model/numfmt.js`, `src/model/style.js`, `src/model/search.js` | slice D |
| `src/ui/grid.js`, `src/ui/app.js`, `src/ui/clipboard.js`, `src/ui/*.css` | slice E |
| `src/edit/edits.js` | slice F |
| `src/main.js` | integration |

Test placement — which file pattern each of the two runners picks up — is stated once, in the
README's *Build and test* section.

## 1. Bytes and the document session

The editor gets the workbook as bytes and returns bytes. `src/ui/document.js` is the only file that
opens an Orca document session; `src/main.js` and `src/ui/app.js` read `globalThis.orca` as well —
the first to tell a missing host from a broken workbook, the second for `orca.files.saveAs` on
export — but neither opens a session:

```js
const doc = await globalThis.orca.document.session({ decode, encode })
doc.value       // Uint8Array, latest accepted bytes
doc.readOnly    // boolean from the host
doc.change({})  // asks the host to encode and save; the argument is ignored — encode() rebuilds
                // from the module's own state, and Orca debounces the call
doc.onReplace((bytes) => { /* outside write landed */ })
doc.flush()     // part of the Orca session API; this plugin never calls it — only the tests do
```

- `decode(content)` never returns `null`. A content that is not `{ kind: 'bytes' }` throws
  `TypeError('a workbook editor needs a bytes document')`. Bytes equal to the run `encode` produced
  are the host accepting the save — `{ bytes }`, and the save mark moves. Any other bytes are a
  document this tab did not write (first open, outside change, reload after a conflict): they return
  `{ bytes, zipStatus }` and the edit log is cleared, because its operations describe a file that no
  longer exists.
- `encode()` takes no argument and returns `{ kind: 'bytes', bytes }`: the last accepted file plus
  the edits since the save mark, or the accepted bytes back unchanged when there is nothing to
  write. It throws when nothing has been read yet, and when the file could not be unpacked — a file
  this editor cannot take apart must not look saved.
- `doc.change(next)` is not used with a value. The three call sites (`commitEdit`, `undo`, `redo`)
  pass `{}` because the host only needs the signal; the new bytes come from `encode`.
- `commitEdit(doc, op)` returns `{ ok: true, op }` — with the settled operation, which may be the
  no-op form — or `{ ok: false, reason }` when the writer would refuse the edit. A refused edit
  never enters the log, so it cannot poison later saves.
- `subscribeDocument({ refused, accepted, discarded })` fills the callbacks; `editState()` reports
  `{ canUndo, canRedo, dirty }`; `setDateSystem(is1904)` tells the writer which date system to use
  when a `Date` is written back as a serial.

Bytes mode has **no merge**: an outside write while edits are unsaved shows Orca's
Reload/Overwrite banner. Do not pretend otherwise.

## 2. `src/write/zip.js` — zip with byte preservation

```js
export const OFFICE_ZIP_MAX_ENTRIES = 4000
export const OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024

/** @typedef {{ name: string, data: Uint8Array, method: number, crc: number,
 *   raw: Uint8Array | null, time: number, date: number, flags: number,
 *   changed?: boolean, uncompressedSize?: number, compressedSize?: number,
 *   versionMadeBy?: number, internalAttributes?: number,
 *   externalAttributes?: number }} ZipEntry */

export function crc32(bytes)                       // -> number
export function readZip(bytes, maxBytes?)          // -> { status: 'ok', entries: ZipEntry[] }
                                                   //  | { status: 'too-large' | 'corrupt', reason: string }
export function writeZip(entries)                  // -> Uint8Array
```

Rules that make the save byte-preserving:

- `raw` is the **exact compressed byte run** read from the original file, and `readZip` sets it on
  every entry it produces, stored or deflated alike. It is `null` only on entries the caller built
  itself.
- `writeZip` copies `raw` verbatim for every entry whose `changed !== true`. Only entries marked
  `changed` are compressed again — with fflate's `deflateSync`, whose unprefixed codecs are raw
  DEFLATE without a zlib or gzip wrapper, exactly what a zip entry carries — and their `crc` must
  be recomputed.
- Reading inflates through fflate's `Inflate`, in 4 KiB steps, so the real output is counted
  against `maxBytes` instead of trusting the declared size.
- Central directory is rebuilt (offsets move), local headers of untouched entries are regenerated
  with the same name, method, time, date, flags and sizes; the payload bytes are the original ones.
- Honour `flags & 0x800` (UTF-8 names). Refuse ZIP64 markers (`0xffffffff`) as `too-large`.
- Never trust declared sizes: inflate and count real bytes against `maxBytes`.

## 3. `src/write/xml.js` — text-level XML surgery

The writer edits the original XML text; it never re-serialises a whole document. Three of the
twelve exports hand back offsets into the string it was given, the other nine hand back strings:

```js
// Offsets: the caller splices, so attribute order, quotes and empty-element form survive.
export function findElement(xml, name)      // -> { name, start, innerStart, innerEnd, end, attrs,
                                            //      selfClosing, tagEnd } | null
export function findElements(xml, name)     // -> array of the same, document order
export function topLevelElements(xml)       // -> array of the same plus `local`; direct children of
                                            //      the root only, so a nested extLst cannot match

// Strings.
export function localName(name)             // -> 'row' for 'x:row'
export function rootName(xml)               // -> root local name, e.g. 'worksheet'; null if none
export function topLevelNames(xml)          // -> string[]
export function getAttr(attrsText, name)    // -> unescaped value | null; a prefixed lookup such as
                                            //      'r:id' falls back to a local-name match
export function setAttr(attrsText, name, value)    // -> attrsText, attribute replaced or appended
export function removeAttr(attrsText, name) // -> attrsText without the attribute and its whitespace
export function replaceElements(xml, name, replacer) // -> new xml; replacer(el, currentText)
                                                     //      returns the replacement or null
export function escapeXml(text)             // -> string, control characters dropped
export function unescapeXml(text)           // -> string, named and numeric entities decoded
```

## 4. `src/write/xlsx-writer.js` — the save that keeps the file

```js
/** @typedef {{ sheet: string, ref: string,
 *   styleFrom?: string,
 *   value: { kind: 'number', value: number }
 *         | { kind: 'text', value: string }
 *         | { kind: 'boolean', value: boolean }
 *         | { kind: 'date', value: number }        // Excel serial day, epoch chosen by the caller
 *         | { kind: 'error', value: string }       // one of Excel's ten error codes
 *         | { kind: 'formula', formula: string, cached?: number | string | boolean | null }
 *         | { kind: 'blank' } }} CellEdit */

/** @typedef {{ kind: 'oplog', ops: CellEdit[] }} EncodeInput */

export function applyEdits(zipEntries, edits)     // -> { status: 'ok', entries: ZipEntry[], notes: string[] }
                                                  //  | { status: 'refused', reason: string }
export function xlsxPackage(zipEntries, diag?)    // -> { entries, files: Map, book } | null
```

Behaviour:

- Writes **only** the target `<c>` elements inside the edited sheets' `<worksheet>` XML. Everything
  else in that sheet (styles, merges, conditional formats, cols, views) keeps its bytes.
- `text` values are written inline (`t="inlineStr"><is><t>`), so `sharedStrings.xml` is untouched.
- A cell that exists keeps its `s` (style) attribute; a new cell in a styled row may copy the style
  of its row's neighbouring cell only when the caller passes `styleFrom` — otherwise no `s`.
- `formula` writes `<f>` and, when `cached` is given, `<v>`.
- When a cell actually changed, the recalculation bookkeeping follows: a `xl/calcChain.xml` the
  package carries is removed together with the relationship and the content-type override that
  pointed at it, and `calcPr fullCalcOnLoad="1"` is written into `xl/workbook.xml` so Excel
  recalculates. Never leave a stale calc chain. An edit that moves no cell leaves the package
  untouched, bookkeeping included.
- Refuse (do not guess) when the sheet part is missing, the XML is not a worksheet, or a
  non-shiftable construct would be corrupted. `notes` explains what was done.
- Must be pure: no DOM, no Orca API.

## 5. `src/model/*` — ExcelJS to a view model

```js
/** @typedef {{ kind: 'number'|'text'|'boolean'|'date'|'error'|'blank',
 *   value: unknown, display: string, formula: string | null, rich?: boolean }} CellValue */

/** @typedef {{ ref: string, row: number, col: number, value: CellValue,
 *   css: { font: object, fill: string | null, border: object, align: object } }} CellView */

/** @typedef {{ name: string, index: number, rowCount: number, colCount: number,
 *   cells: Map<string, CellView>,                  // key `r,c`, 0-based
 *   colWidths: Map<number, number>,                // px
 *   rowHeights: Map<number, number>,               // px
 *   merges: Array<{ r1: number, c1: number, r2: number, c2: number }>,
 *   freeze: { rows: number, cols: number },
 *   hidden: { rows: Set<number>, cols: Set<number> } }} SheetModel */

/** @typedef {{ sheets: SheetModel[], active: number, date1904: boolean,
 *   notes: string[] }} WorkbookModel */

export async function loadWorkbook(bytes)         // -> WorkbookModel
export function normalizeAbsoluteTargets(bytes)   // -> Uint8Array | null
```

`loadWorkbook` parses a **reader-only copy** built in memory: XML namespace prefixes and their
declarations are stripped from `xl/workbook.xml` and worksheet XML (comment bodies are protected),
`xl/drawings/*` except `.vml` and `xl/tables/*` are removed from the archive copy, `<drawing>` and
`<tableParts>` references are removed from worksheet XML, and worksheet relationship targets are
normalized to paths relative to the sheet (absolute targets and `.`/`..` segments included). The
copy exists only for ExcelJS. `src/ui/document.js` keeps the accepted bytes' own ZIP entries, so a
normalized archive is never the base of a save.

An archive that cannot be read fails with `workbook archive cannot be read safely (<status>)`.
`loadWorkbook` also compares ExcelJS's worksheet count with the `<sheet>` entries in
`xl/workbook.xml` and throws when they differ, so a partially loaded workbook never looks complete.

`normalizeAbsoluteTargets` is the exported relationship-only helper (absolute and dot-segment
`Target` values → sheet-relative paths) used by the model tests; `loadWorkbook` does not route
through it.

`WorkbookModel` also carries `date1904: boolean` — the file's date epoch. `formatValue` takes it as
an optional fourth argument (`formatValue(value, numFmt, isDate, date1904)`), and
`src/ui/document.js` uses it to write a `Date` back as the serial number the file expects.

`notes` is what the read had to leave out or approximate, de-duplicated across sheets: the
conditional-formatting sentence (`Conditional formatting rules are preserved, but are not
displayed.`) and unresolved style values (theme or indexed colours left unset, a non-solid pattern
fill drawn solid). `src/ui/app.js` renders `model.notes` in the notice above the grid, next to the
read-only warning. A workbook ExcelJS still cannot parse fails with Orca's error, so a warning is
not a claim that every workbook opens.

- Pure: no DOM. Runs under `node --test`.
- Styles resolve to CSS-ready values: `fill` is a hex string like `#1F4E79` or `null`;
  `font` is `{ bold, italic, underline, size, color, family, strike, vertAlign }`;
  `border` is `{ top, right, bottom, left }` with `{ style, color }`;
  `align` is `{ horizontal, vertical, wrapText, indent, textRotation }`.
- `display` is the formatted number/date text per `numFmt` (subset in `numfmt.js`), so the grid
  shows `test/fixtures/styled.xlsx` as Excel does: `2.00%` in `F2`, `200,274.00` in `B2`.
- Column width in characters (`wch`) converts to px with the Excel rule `wch * 7 + 5`, clamped;
  row height `pt` to px at `pt * 4 / 3`.
- `freeze` comes from `views[0].state === 'frozen'`.

## 6. `src/ui/*` — the grid

`grid.js` exports a class or factory that renders a `SheetModel` into a container with
virtualisation, sticky headers honouring `freeze`, selection and keyboard navigation. No Orca API
inside `grid.js`; `app.js` owns the chrome and calls into the session it was handed. `app.js` also
wires the toolbar: copy (TSV, CSV, Markdown, formatted HTML, "формулы как текст") through
`clipboard.js`, export of the whole sheet as CSV or Markdown through `orca.files.saveAs`, workbook
search through `src/model/search.js`, and row/column insertion and deletion through
`commitStructural`. Search cursors carry a monotonic `revision`; every accepted edit, paste,
undo/redo and model replacement increments it, so a continuation can never replay an old value.
`grid.setFormulaResultsStale(boolean)` is set after successful edits and history moves; the grid
never claims a formula was recalculated.

A structural commit returns the bytes of the staged package; the app reloads the whole model from
`currentDocumentBytes()` instead of patching the grid. A refusal leaves the model, history and file
untouched and is shown as a message.

## 7. `src/edit/edits.js` — operations and undo

```js
export function createEditLog(options?)   // -> { apply(op), undo(), redo(), ops(), opsSince(mark),
                                          //      markSaved(), clear(), checkpoint(), restore(),
                                          //      canUndo, canRedo, undoSize, dropped }
```

Cell ops are `{ type: 'cell', sheet, ref, before: CellValue, after: CellValue }` with inverses kept
for undo. Structural ops are `{ type: 'structural', sheet, axis: 'row'|'column',
action: 'insert'|'delete', index, count, beforeEntries, afterEntries }`: the two entry arrays are
the package snapshots before and after the whole transaction, because a structural edit changes
addresses, formulas and related OOXML beyond one cell. Undo of a structural record appends
`{ type: 'restore', entries: beforeEntries }` (or `afterEntries` on redo) so the file can be
returned to either snapshot even after the host accepted a save. No DOM, no Orca API.

`commitStructural(doc, operation)` in `document.js` preflights the complete staged package with
`applyStructural` before recording anything: a refusal changes neither the log nor the bytes. While
a save is pending or the document is read-only, structural commits are refused. An outside
replacement clears the history and its snapshots. `currentDocumentBytes()` returns the staged ZIP,
which the UI reloads after a structural commit or history move; the model cannot be patched
cell-by-cell when addresses and related objects moved.

Settled behaviour: `apply` returns `null` for a no-op (`after` equals `before`) and clears redo
otherwise; consecutive edits in one cell coalesce inside `coalesceWindowMs` (600) so one Ctrl+Z
rewinds a whole burst; the undo stack keeps the last 1000 operations while `ops()` keeps every
edit; `opsSince(mark)` returns a collapsed per-cell delta (one op per touched cell, net-zero
dropped), which is exactly what the writer needs; `markSaved()` returns a mark whose seq survives
`clear()`.

**A save closes the coalescing batch.** `markSaved`, `undo`, `redo` and `clear` all close it, so the
next edit starts a new operation. An `undo` of an operation at or below the save mark does not
vanish from the log: it appends the inverse operation, because the edit it reverses may already be
in the file. `redo` likewise re-appends with a fresh seq. Without this, a save followed by an undo
produced an empty delta and left the file changed.

## 8. The worker and the hand-off

`worker/main.js` is the plugin's Node process (`trusted-worker`, `main` in the manifest). It is the
only place that may call `agents.list` and `terminal.sendText`, because both need capabilities a
frame cannot hold:

```js
export default function activate(orca)         // required; a named export is never called
orca.panels.handle('agent.list', handler)      // -> { agents: Agent[] }, already ranked
orca.panels.handle('agent.send', handler)      // { terminalId, text, enter? }
                                               // -> { accepted, sent, unknown, reason? }
```

`worker/handoff.js` holds the pure part — `buildHandoff` (the context header plus a Markdown or
fenced TSV body; a TSV fence is at least three backticks and one longer than any run at the start
of a body line), `fitHandoffRows` (leading row costs against a character budget), `rankAgents`,
`readSend`, and `TEXT_LIMIT` 4096 — so the frame and the worker share the framing and validation
without Orca.

`readSend` validates before any host call: a non-empty `terminalId` (at most 1024 characters), a
non-empty `text`, and a boolean `enter` when present. `agent.send` then refuses text longer than
`TEXT_LIMIT` without calling the host. The hand-off is never split or truncated: Orca offers no
atomicity or idempotency across calls. For text over the message limit the frame disables sending
and offers a button to copy the full hand-off for manual pasting. The frame separately limits the
selection to 200,000 cells: beyond that it refuses both sending and copying. Before materializing a
copy, `selectionFitsTextBudget` conservatively estimates its size against 8 × 1024 × 1024 UTF-16
code units, reserving the sheet name length, range length and 128 units for the header. This bound
can reject a smaller actual output; exceeding either copy limit copies nothing and asks the person
to reduce the selection. One send per terminal is allowed at a time; a second request for the same
terminal while the first is in flight is refused.

A valid request makes exactly one `terminal.sendText` call with `enter: true` only when the person
asked for it. `sent` is that one message. On acceptance the answer is
`{ accepted: true, sent, unknown: false }`; a non-accepting host result is
`{ accepted: false, sent, unknown: false, reason }`. A thrown host call returns
`{ accepted: false, sent, unknown, reason }`: `unknown` is false only for codes that prove the host
refused before typing (`invalid_request`, `invalid_params`, `unknown_method`, `capability_denied`,
`consent_required`, `not_permitted`, `rate_limited`, `too_large`), and true otherwise — a timeout
may already have typed the text. The worker never retries. The frame calls the handlers with
`orca.plugin.request`; nothing is sent without a press.

## 9. `src/ui/clipboard.js` — clipboard and export text

```js
export function selectionToTsv(sheet, selection, { formulas = false })       // -> string
export function selectionToCsv(sheet, selection, { formulas = false })       // -> string, RFC 4180, CRLF
export function selectionToMarkdown(sheet, selection, { formulas = false })  // -> string, GFM table
export function selectionToHtml(sheet, selection, { formulas = false, maxBytes }) // -> string
export function copyText(text, target = globalThis.document)                 // -> boolean
export function copyHtml(html, text, target = globalThis.document)            // -> boolean
```

`selection` is `{ r1, c1, r2, c2 }`, 0-based and inclusive. The formatters are pure — no DOM,
no Orca API — so `node --test` runs them directly; they copy the display text, or `=FORMULA` when
`formulas` is set. Text that looks like a formula but is not one is defused in every text format.
`selectionToHtml` emits an escaped HTML table with inline font, fill, alignment and border styles,
`rowspan`/`colspan` for merged cells, and refuses (throws `RangeError`) above 250,000 cells, above
8 MiB of serialized UTF-8 (or a supplied `maxBytes`), or when a clipped merge would exceed the
browser's `colspan`/`rowspan` ceilings, rather than dropping cells silently. `copyText` and
`copyHtml` are the file's DOM-facing functions: the isolated frame has no `navigator.clipboard`, so
they select a hidden textarea, run the legacy `copy` command, restore focus and report whether the
copy happened. `copyHtml` puts `text/html` and the plain fallback on the same clipboard event.

## 10. Test fixtures

The three workbooks the tests read — `rich.xlsx`, `macro.xlsm`, `styled.xlsx` — are **not** in the
repository. A public plugin repository carries no spreadsheet binaries (the Orca catalog rejects
them as `binary-files`), so they are assembled at test time by `test/helpers/build-fixtures.mjs`:

- `buildFixtures()` returns `{ 'rich.xlsx', 'macro.xlsm', 'styled.xlsx' }` as `Uint8Array`s, with no
  filesystem, clock or randomness involved — two runs produce byte-identical archives.
- `node test/helpers/build-fixtures.mjs` (or `npm run fixtures`) writes the three files into
  `test/fixtures/`, which is git-ignored (`/test/fixtures/*.xlsx`, `*.xlsm`, `*.bin`, `*.tmp`).
  Deleting that directory is safe: the next test run rebuilds it.
- `test/helpers/fixtures.mjs` exports `fixturePath(name)` and `fixtureBytes(name)`. The first call in
  a test process builds the books and writes them, so `git clone && npm ci && npm test` needs no
  preparation and no Python.

The parts are plain OOXML strings packed by fflate. That keeps the fixtures independent of both
things under test — the reader (`src/model/workbook.js` over ExcelJS) and the writer
(`src/write/zip.js`) — since no ExcelJS workbook is used to build them, a reader bug cannot hide
behind a fixture that was written by that reader.

What the books contain, and which parts of this contract they exercise:

| Book | Contents |
| --- | --- |
| `rich.xlsx` | `Summary` / `Notes` / `Plain`. Painted header (`FF1F4E79` fill, bold white, centered, wrapped, thin `FFB0B0B0` borders, 30 pt row), widths 44/13/20/10/14, five data rows with `#,##0`, `#,##0.00 "RUB"`, `0.00%` and a text date column, `=B{r}*48.7` formulas with no cached result, banded rows, a bold `Total` row with `=SUM(...)`, the merged `A9:E9`, freeze `B2`, conditional formatting on `D2:D6`, the defined name `SalesTotal`, a 200-row `Plain` sheet, and a `Notes` sheet with no styling at all. |
| `macro.xlsm` | The same parts plus `xl/vbaProject.bin` (10 240 deterministic bytes) and its `[Content_Types].xml` override — the byte-preserving save must copy it untouched. |
| `styled.xlsx` | One `Ledger` sheet, 166 × 16, the same painted header and thin borders on every cell, banded body rows, widths 43/14, `#,##0.00` / `#,##0` / `#,##0;[Red]-#,##0` / `0.00%` / `#,##0.0`, formulas in `C` and `E`, freeze `B2`. |

Tests assert on properties, not on recorded CRC or size numbers: an assertion that pinned bytes would
only restate the generator. `test/zip.spec.mjs` checks instead that each header checksum equals
`crc32(data)`, that no entry is empty, and that the part count and names are the ones the fixture is
supposed to have.

