# Changelog

## 1.0.0

First public release.

- Custom editor for `.xlsx` and `.xlsm` that draws what the file carries: fills, fonts, borders,
  alignment, number formats, merged cells, frozen panes, column widths and row heights, on every
  sheet.
- Editing: type a value, overwrite a cell, clear a selection, paste a block of text, undo and redo.
  Numbers, text, booleans, formulas, dates and error values round-trip with their type.
- Saving rewrites only the cells that changed inside the sheets that changed. Every other part of
  the package — styles, shared strings, drawings, charts, macros, pivot caches — keeps its exact
  compressed bytes, so a workbook with a macro project opens and saves without touching it.
- Copy the selection as TSV, CSV or a Markdown table, and export a whole sheet to a file.
- Read-only documents, outside writes and refusals by the writer are reported instead of being
  quietly ignored; an edit the writer cannot take never blocks the rest of the save.
