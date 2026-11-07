<div align="center">

[Русская версия](README.md)

<img src="./banner.png" alt="XLSX Editor — Excel workbook editor for Orca with byte-preserving saves" width="100%">

<p>
  <a href="https://github.com/narimanisakhanov-creator/orca-xlsx-editor/releases/latest"><img src="https://img.shields.io/github/v/release/narimanisakhanov-creator/orca-xlsx-editor?label=release&color=5f7f95" alt="Latest release"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-6b7280" alt="MIT license"></a>
  <img src="https://img.shields.io/badge/Orca-%E2%89%A51.4.248-58758a" alt="Orca 1.4.248 or newer">
  <img src="https://img.shields.io/badge/tests-npm%20test-6f8796" alt="Tests: npm test">
  <img src="https://img.shields.io/badge/formats-xlsx%20%2F%20xlsm-4f7c8a" alt="xlsx and xlsm formats">
</p>

<p><b>An Excel workbook editor for Orca: file-native formatting, focused cell edits and untouched ZIP parts preserved on save.</b></p>

<p>
  <a href="#installation">Installation</a> ·
  <a href="#capabilities">Capabilities</a> ·
  <a href="#hand-off-to-an-agent">Hand-off to an agent</a> ·
  <a href="#access-and-security">Access</a> ·
  <a href="#how-saving-works">How saving works</a> ·
  <a href="#limitations">Limitations</a> ·
  <a href="#development">Development</a> ·
  <a href="#changes">Changes</a>
</p>

</div>

---

![Workbook with fills, borders, number formats and a frozen row](docs/images/styled-workbook.png)

<p align="center"><sub>A real Orca capture from the synthetic workbook <code>test/fixtures/styled.xlsx</code>; see <a href="docs/screenshots.md">docs/screenshots.md</a> for how it was produced.</sub></p>

## Installation

**From Git.** Open this link in Orca:

<a href="orca://plugins/narimanisakhanov-creator.xlsx-editor?catalog=https://github.com/narimanisakhanov-creator/orca-xlsx-editor.git">Install from Git</a>

**Development mode.** From the root of a cloned repository, run:

```sh
orca plugin dev load .
```

Requires Orca `>=1.4.248`. Node.js `>=20.6` is required for builds and tests.

## Capabilities

- Opens `.xlsx` and `.xlsm` workbooks, including every sheet.
- Draws fills, fonts, borders, alignment, number formats, merged cells, frozen panes, column widths and row heights.
- Edits cell values and types, with formulas, clearing, block paste, undo and redo.
- Inserts and deletes whole rows and columns. Formulas on every sheet, defined names, merged ranges, tables, drawings, filters and column geometry are translated in one transaction; an unsupported reference or object refuses the whole operation. Undo and redo work even after a save.
- Searches values and formulas across all sheets (`Найти`, Ctrl+F), navigating to a match and continuing an incomplete page explicitly.
- Copies a selection as TSV, CSV, Markdown or formatted HTML, and exports a complete sheet.
- After an edit, undo or redo, marks formulas and says that cached results may be out of date: the editor never recalculates.
- Hands the table to an agent's terminal in Orca — with a preview; text over the message limit is not sent.
- Leaves macros, drawings, charts and pivot caches untouched on save; objects the editor does not draw still survive.
- Works locally without network access.

## Hand-off to an agent

The **В агента** button in the editor toolbar turns the selection (or the active cell) into a
Markdown table or TSV and types it into the terminal of a chosen agent in the current workspace.
Nothing is sent without a press; by default the text is inserted **without** Enter so it can be
read and extended, and the "Send immediately (Enter)" checkbox sends that same message with Enter.

For a sendable selection, the preview is exactly the text that will be sent. A hand-off is **one**
`terminal.sendText` message of at most 4096 characters (Orca's limit). Longer text is neither
split nor truncated: sending is refused and nothing is sent. **Скопировать полный текст** offers
copying for manual pasting. A selection over 200,000 cells cannot be sent or copied this way.
Copying also uses a conservative preflight estimate with a limit of 8 × 1024 × 1024 UTF-16 code
units, including an allowance for the header; it can refuse text whose actual size would be smaller.
When a limit is exceeded nothing is copied; reduce the selection. The plugin does not retry by
itself; after an unknown outcome (a timeout, for example), inspect the terminal before sending again.

For this the plugin runs in the `trusted-worker` profile and asks for two capabilities when enabled:
`agent:read` to see the workspace's agents, and `terminal:send` to type into one terminal. The
hand-off adds no new capability.

## Access and security

The editor is sandboxed: it draws a workbook, receives bytes and hands new bytes back. It never learns the file's path, and has no file system and no network.

The **В агента** hand-off needs two Orca APIs a frame cannot call, so the plugin ships a Node worker and asks for two capabilities:

- `agent:read` — `agents.list`, to offer the agents of the current workspace;
- `terminal:send` — `terminal.sendText`, to type the chosen text into the chosen terminal.

Both calls live only in `worker/main.js`, only inside its two handlers, and only after a press. The
text is inserted without Enter by default; Enter goes with the message only when the checkbox is
ticked. A hand-off is one message within Orca's 4096-character limit, and the plugin never sends
anything by itself — no timer, no event. Details: [SECURITY.md](SECURITY.md).

## How saving works

An Excel workbook is a ZIP of XML parts; the editor changes only `<c>` elements in the XML of sheets that were edited. Untouched ZIP entries retain their original compressed bytes, while text values are written inline so `sharedStrings.xml` is not rewritten. Macros and every other untouched part remain byte-for-byte unchanged. When a cell actually changes and the workbook carries `calcChain.xml`, the chain is removed, relationships are updated, and `fullCalcOnLoad="1"` asks Excel to recalculate; a save with no cell changes leaves this bookkeeping untouched.

To open more real files, ExcelJS receives a **separate reading-only copy**: XML namespace prefixes and their declarations are stripped, `xl/drawings/*` (except `.vml`) and `xl/tables/*` parts are removed, `<drawing>` and `<tableParts>` references are removed from the sheet XML, and sheet relationship targets are normalized to paths relative to the sheet (absolute targets and `.`/`..` segments included). That copy lives in memory and is never written; saving starts from the original ZIP entries of the accepted file plus the cell edits. When a sheet carries conditional formatting, a warning above the grid says the rules are preserved but not displayed. This does not mean every workbook now opens: a file ExcelJS still cannot parse for another reason still fails, and the editor shows the error.

## Limitations

- Charts, pivot tables, sparklines and other drawing objects are not rendered; they remain in the workbook on save.
- Formulas are not evaluated: the editor shows the cached result or formula text, and an edited formula is left for Excel to recalculate on open.
- Row and column insertion/deletion works only where dependencies can be rewritten safely. A workbook with sheet protection, extensions, an unsupported formula (for example `INDIRECT`, `OFFSET`, 3D or external references), a chart or an unresolvable object relationship is refused as a whole: there is no partial edit. A structural edit removes the calculation chain and asks Excel to recalculate on open.
- Formatted copying is limited to 250,000 cells, 8 MiB of HTML and the browser ceilings for merged cells; when it refuses, the same data can be copied as TSV.
- Only `.xlsx` and `.xlsm` are supported; `.xls`, `.xlsb` and `.ods` are not.
- Conditional formatting is preserved but not displayed; a workbook that carries it shows a warning above the grid.
- The reading compatibility pass covers ExcelJS's known limits on an in-memory copy: namespace prefixes and declarations, `xl/drawings/*` (except `.vml`) and `xl/tables/*` parts and their references, and relationship targets with absolute or dot-segment paths. A workbook ExcelJS still cannot parse for another reason does not open; there is no guarantee that every workbook opens.
- There is no network access: the editor does not load external resources.
- The hand-off works in the desktop app only: in the browser client the frame has no worker.

## Development

```sh
npm ci
npm run build
npm test
npm run check:dist
orca plugin validate .
orca plugin check .
orca plugin test .
```

`npm test` runs the Node.js tests. `npm run build` creates `editor/assets/` and checks the resulting distribution; `check:dist` repeats those checks without writing. `orca plugin validate`, `check` and `test` run Orca's manifest, package and test checks.

The test workbooks are not in the repository: a public plugin repository carries no binary files, which is what the Orca catalog requires. `rich.xlsx`, `macro.xlsm` and `styled.xlsx` are assembled by `test/helpers/build-fixtures.mjs` from plain OOXML strings, packed into a ZIP by `fflate`; a test gets a book through `test/helpers/fixtures.mjs`, which writes the files on first use. So `git clone && npm ci && npm test` works with no preparation and no Python, and the build is deterministic: zip entries carry a fixed timestamp, and two runs produce byte-identical workbooks. To rebuild the books by hand, run `node test/helpers/build-fixtures.mjs`; the `test/fixtures/` directory is git-ignored.

Release order: `orca plugin release-entry` reads the **committed HEAD**, not the working tree, and `orca-catalog.json` is part of the hashed release content. Build and verify all plugin content, including the existing catalog, commit it as **commit A**, and tag A. Then, on a clean tree with A checked out, run `node scripts/write-catalog.mjs`: the new entry receives A's `commit` and `contentHash`. Commit the generated catalog separately as **commit B**. The catalog bytes present in A are included in A's hash; do not amend A with the generated entry, move the tag, or re-pin the entry to B, which would create a self-reference.

## Changes

- [Changelog](CHANGELOG.md)
- [Security and access](SECURITY.md)
- [MIT license](LICENSE)
- [Third-party notices](THIRD_PARTY_NOTICES.md)
- [Русская версия](README.md)
