# Module contract

The interfaces between the slices of this plugin, as the code implements them.

Runtime: browser editor frame (`frame: "isolated"`), bundled by esbuild to `editor/assets/`.
No network, no `eval` / `new Function`, no external requests. Everything is relative to the plugin.

## Slices and the files they own

| Path | Owner |
| --- | --- |
| `build/build.mjs`, `package.json`, `editor/index.html`, `editor/boot.js`, `editor/page.css` | slice A |
| `src/write/zip.js`, `src/write/zip-budget.js` | slice B |
| `src/write/xml.js`, `src/write/xlsx-writer.js` | slice C |
| `src/model/workbook.js`, `src/model/numfmt.js`, `src/model/style.js` | slice D |
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

/** @typedef {{ sheets: SheetModel[], active: number }} WorkbookModel */

export async function loadWorkbook(bytes)         // -> WorkbookModel
export function normalizeAbsoluteTargets(bytes)   // -> Uint8Array | null
```

`normalizeAbsoluteTargets` exists because ExcelJS 4.4.0 crashes on a worksheet relationship whose
`Target` is an absolute OPC path (`/xl/tables/table1.xml`), which real writers produce. It rewrites
those targets to paths relative to the sheet and returns a rebuilt archive, or `null` when the file
needs no change; `loadWorkbook` reads the copy and never mutates the caller's bytes.

`WorkbookModel` also carries `date1904: boolean` — the file's date epoch. `formatValue` takes it as
an optional fourth argument (`formatValue(value, numFmt, isDate, date1904)`), and
`src/ui/document.js` uses it to write a `Date` back as the serial number the file expects.

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
wires the two toolbar menus — copy (TSV, Markdown, CSV, "формулы как текст") through
`clipboard.js`, export of the whole sheet as CSV or Markdown through `orca.files.saveAs`.

## 7. `src/edit/edits.js` — operations and undo

```js
export function createEditLog(options?)   // -> { apply(op), undo(), redo(), ops(), opsSince(mark),
                                          //      markSaved(), clear(), canUndo, canRedo, undoSize, dropped }
```

Ops are `{ type: 'cell', sheet, ref, before: CellValue, after: CellValue }` with inverses kept for
undo. No DOM, no Orca API.

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
orca.panels.handle('agent.send', handler)      // { terminalId, text, enter? } -> { accepted, sent, truncated, reason? }
```

`worker/handoff.js` holds the pure part — `buildHandoff`, `truncateText` (the host refuses a
`terminal.sendText` over 4096 characters), `rankAgents`, `readSend` — so the frame can build the
same text it will send and both sides can be tested without Orca. The frame calls the handlers with
`orca.plugin.request`; nothing is sent without a press.

## 9. `src/ui/clipboard.js` — clipboard and export text

```js
export function selectionToTsv(sheet, selection, { formulas = false })       // -> string
export function selectionToCsv(sheet, selection, { formulas = false })       // -> string, RFC 4180, CRLF
export function selectionToMarkdown(sheet, selection, { formulas = false })  // -> string, GFM table
export function copyText(text, target = globalThis.document)                 // -> boolean
```

`selection` is `{ r1, c1, r2, c2 }`, 0-based and inclusive. The three formatters are pure — no DOM,
no Orca API — so `node --test` runs them directly; they copy the display text, or `=FORMULA` when
`formulas` is set. `copyText` is the file's only DOM-facing function: the isolated frame has no
`navigator.clipboard`, so it selects a hidden textarea and runs the legacy `copy` command, then
removes the node and restores focus. It returns whether the copy happened.
