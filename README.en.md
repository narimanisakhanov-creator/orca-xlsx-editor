<div align="center">

[Русская версия](README.md)

<img src="./banner.png" alt="XLSX Editor — an Excel workbook editor in Orca" width="100%">

<p>
  <a href="https://github.com/narimanisakhanov-creator/orca-xlsx-editor/releases/latest"><img src="https://img.shields.io/github/v/release/narimanisakhanov-creator/orca-xlsx-editor?label=release&color=5f7f95" alt="Latest release"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-6b7280" alt="MIT license"></a>
  <img src="https://img.shields.io/badge/Orca-%E2%89%A51.4.248-58758a" alt="Orca 1.4.248 or newer">
  <img src="https://img.shields.io/badge/formats-xlsx%20%2F%20xlsm-4f7c8a" alt="xlsx and xlsm formats">
</p>

**An Excel workbook editor in Orca: the file's formatting, focused edits, and every untouched byte preserved.**

</div>

---

![A workbook with fills, borders, number formats and a frozen row](docs/images/styled-workbook.png)

## Install

[Open in Orca](orca://plugins/narimanisakhanov-creator.xlsx-editor?catalog=https://github.com/narimanisakhanov-creator/orca-xlsx-editor.git) — or add the community catalog `https://github.com/stablyai/orca-community-plugins.git`, ref `main`, and install **XLSX Editor** from the list.

From a clone, for development:

```sh
npm ci
npm run build
orca plugin dev load .
```

Requires Orca `>=1.4.248` and Node.js `>=20.6`.

## What it does

- Opens `.xlsx` and `.xlsm`, every sheet: fills, fonts, borders, alignment, number formats, merged cells, frozen panes, column widths and row heights.
- Edits values and formulas, clears cells, pastes a block, undoes and redoes.
- Inserts and deletes rows or columns. Formulas on every sheet, defined names, merges, tables, drawings, filters and relationships are rewritten in one transaction; an unsupported construct refuses the whole edit. Undo and redo survive an accepted save.
- Searches values and formulas across all sheets (`Ctrl+F`), jumps to a match and continues an incomplete page.
- Copies a selection as TSV, CSV, Markdown or formatted HTML; exports a sheet to CSV or Markdown.
- Marks formulas after edits and says cached results may be stale: the editor never recalculates.
- Hands a table to an agent's terminal in Orca — with a preview and one message of at most 4096 characters.

## How it works

**Reading.** ExcelJS receives a separate in-memory copy prepared for it: namespaces, CDATA and relationships are normalized and parts it cannot parse are handled safely. The original file is not modified.

**Saving.** Only the touched XML elements change; every other ZIP entry keeps its original compressed bytes. Text values are written inline, so `sharedStrings.xml` is not rewritten. Macros, drawings, charts and pivot caches stay untouched. When a change needs recalculation, `calcChain.xml` is removed and `fullCalcOnLoad="1"` is set.

**Structural edits.** Rows and columns change transactionally: the whole package and its related parts are checked first, then rewritten and saved. An unsupported formula, a protected sheet or an unresolved relationship refuses the edit with no partial changes. History keeps package snapshots, so undo works even after a save.

**Hand-off.** The **To agent** button turns the selection into Markdown or TSV and types it with one `terminal.sendText` call. Text longer than 4096 characters is not sent or trimmed — copy the full text manually instead. Enter is off by default. It adds no capabilities beyond `agent:read` and `terminal:send`.

**Language.** The interface follows Orca's UI language: `ru` gives Russian, everything else English.

## Limitations

- Formulas are not evaluated; the editor shows the cached result or the formula text.
- Charts, pivot tables and other drawing objects are not rendered, but are preserved.
- Structural edits refuse protected sheets, extensions, charts, `INDIRECT`/`OFFSET`, 3D and external references.
- Formatted copying is limited to 250,000 cells, 8 MiB of HTML and browser limits for merges; TSV remains available.
- Conditional formatting is preserved but not displayed; a warning appears above the grid.
- Only `.xlsx` and `.xlsm` are supported; `.xls`, `.xlsb` and `.ods` are not. No network access.
- The hand-off works in the Orca desktop app only.

## Development

```sh
npm ci
npm test
npm run build
npm run check:dist
orca plugin validate .
orca plugin check .
orca plugin test .
```

`npm test` builds the test workbooks from OOXML strings in the git-ignored `test/fixtures` directory and runs the full Node test suite. `npm run build` regenerates `editor/assets/`; `check:dist` verifies the build without writing. Release: commit the content, create an annotated tag, then write `orca-catalog.json` in a separate commit (the order is documented in `scripts/write-catalog.mjs`).

- [Changelog](CHANGELOG.md)
- [Security and access](SECURITY.md)
- [Module contract](docs/CONTRACT.md)
- [MIT license](LICENSE)
- [Third-party notices](THIRD_PARTY_NOTICES.md)
- [Русская версия](README.md)
