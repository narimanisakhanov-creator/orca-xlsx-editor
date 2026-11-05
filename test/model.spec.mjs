// Contract §5: ExcelJS to the view model. Runs under `node --test`, no DOM, no Orca API.
//
// The fixtures are read-only and are built by test/helpers/build-fixtures.mjs from hand-written
// OOXML parts, so the expected values here are what an independent writer produced.

import assert from 'node:assert/strict'
import { existsSync, readFileSync } from 'node:fs'
import { test } from 'node:test'

import ExcelJS from 'exceljs'
import { unzipSync, zipSync } from 'fflate'

import { loadWorkbook, normalizeAbsoluteTargets } from '../src/model/workbook.js'
import { formatValue, isDateFormat } from '../src/model/numfmt.js'
import { resolveStyle, themePalette } from '../src/model/style.js'
import { fixtureBytes as bytesOf } from './helpers/fixtures.mjs'

async function loadFixture(name) {
  return loadWorkbook(bytesOf(name))
}

/* ------------------------------------------------------------------ rich.xlsx ------------------ */

test('rich.xlsx: three sheets, in file order, first one active', async () => {
  const model = await loadFixture('rich.xlsx')

  assert.deepEqual(model.sheets.map((sheet) => sheet.name), ['Summary', 'Notes', 'Plain'])
  assert.deepEqual(model.sheets.map((sheet) => sheet.index), [0, 1, 2])
  assert.equal(model.active, 0)
  assert.equal(model.date1904, false)
  assert.deepEqual(model.notes, [])
})

test('rich.xlsx Summary: header A1 font, fill and alignment', async () => {
  const [summary] = (await loadFixture('rich.xlsx')).sheets
  const a1 = summary.cells.get('0,0')

  assert.equal(a1.ref, 'A1')
  assert.equal(a1.row, 0)
  assert.equal(a1.col, 0)
  assert.equal(a1.value.kind, 'text')
  assert.equal(a1.value.display, 'Store')

  assert.equal(a1.css.font.bold, true)
  assert.equal(a1.css.font.size, 11)
  assert.equal(a1.css.font.color, '#FFFFFF')
  assert.equal(a1.css.fill, '#1F4E79')
  // ExcelJS spells a vertically centred cell "middle"; the model hands the grid CSS's "center".
  assert.equal(a1.css.align.horizontal, 'center')
  assert.equal(a1.css.align.vertical, 'center')
  assert.equal(a1.css.align.wrapText, true)
  assert.deepEqual(a1.css.border.top, { style: 'thin', color: '#B0B0B0' })
  assert.deepEqual(a1.css.border.left, { style: 'thin', color: '#B0B0B0' })
})

test('rich.xlsx Summary: freeze pane, column widths, row height', async () => {
  const [summary] = (await loadFixture('rich.xlsx')).sheets

  assert.deepEqual(summary.freeze, { rows: 1, cols: 1 })
  // The fixture writes width 44 for column A; the Excel rule wch * 7 + 5 gives 313 px.
  assert.equal(summary.colWidths.get(0), 313)
  assert.equal(summary.colWidths.get(1), 13 * 7 + 5)
  assert.equal(summary.colWidths.get(2), 20 * 7 + 5)
  assert.equal(summary.colWidths.get(3), 10 * 7 + 5)
  assert.equal(summary.colWidths.get(4), 14 * 7 + 5)
  // 30 pt at pt * 4 / 3.
  assert.equal(summary.rowHeights.get(0), 40)
  assert.equal(summary.rowHeights.has(1), false)
})

test('rich.xlsx Summary: row/col counts and the merged range', async () => {
  const [summary, notes, plain] = (await loadFixture('rich.xlsx')).sheets

  assert.equal(summary.rowCount, 9)
  assert.equal(summary.colCount, 5)
  assert.deepEqual(summary.merges, [{ r1: 8, c1: 0, r2: 8, c2: 4 }])
  // The merged block's own cell survives; the covered cells hold no value of their own.
  assert.equal(summary.cells.get('8,0').value.value, 'Prepared from the store ledger')

  assert.deepEqual(notes.merges, [])
  assert.equal(plain.rowCount, 200)
  assert.equal(plain.colCount, 2)
})

test('rich.xlsx Summary: number, date and formula cells render as Excel does', async () => {
  const [summary] = (await loadFixture('rich.xlsx')).sheets

  const margin = summary.cells.get('1,3')
  assert.equal(margin.value.kind, 'number')
  assert.equal(margin.value.value, 0.0812)
  assert.equal(margin.value.display, '8.12%')

  const opened = summary.cells.get('1,4')
  // The fixture stores the date column as text, so the format does not turn it into a date.
  assert.equal(opened.value.kind, 'text')
  assert.equal(opened.value.display, '2019-04-01')

  const sales = summary.cells.get('1,1')
  assert.equal(sales.value.display, '282,349')
  assert.equal(sales.css.align.horizontal, 'right')

  const revenue = summary.cells.get('1,2')
  assert.equal(revenue.value.formula, 'B2*48.7')
  // rich.xlsx carries `<f>` with an empty `<v/>`: no cached result, so the formula is the text.
  assert.equal(revenue.value.display, 'B2*48.7')
})

/* ------------------------------------------------------------------ cached formula ------------ */

test('formula: a cached result is formatted with the cell number format', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Calc')
  sheet.getCell('B2').value = 282349
  const revenue = sheet.getCell('C2')
  revenue.value = { formula: 'B2*48.7', result: 13750396.3 }
  revenue.numFmt = '#,##0.00 "RUB"'

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const cell = model.sheets[0].cells.get('1,2')

  assert.equal(cell.value.kind, 'number')
  assert.equal(cell.value.value, 13750396.3)
  assert.equal(cell.value.formula, 'B2*48.7')
  assert.equal(cell.value.display, '13,750,396.30 RUB')
})

/* ------------------------------------------------------------------ date system ---------------- */

test('date1904: the 1904 date system is reported and used for reading dates', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.properties.date1904 = true
  const sheet = workbook.addWorksheet('Dated')
  const cell = sheet.getCell('A1')
  cell.value = new Date(Date.UTC(2024, 0, 15))
  cell.numFmt = 'yyyy-mm-dd'

  // A 1904 book stores 2024-01-15 as 43844; a 1900 book would store 45306. The date is only read
  // back correctly when the model's flag agrees with the file's epoch.
  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  assert.equal(model.date1904, true)

  const loaded = model.sheets[0].cells.get('0,0')
  assert.equal(loaded.value.kind, 'date')
  assert.equal(loaded.value.value.toISOString(), '2024-01-15T00:00:00.000Z')
  assert.equal(loaded.value.display, '2024-01-15')
})

/* ------------------------------------------------------------------ styled.xlsx ---------------- */

test('styled.xlsx: 166x16 Ledger sheet, freeze B2, banded fill, thin borders', async () => {
  const model = await loadFixture('styled.xlsx')
  assert.equal(model.sheets.length, 1)

  const [sheet] = model.sheets
  assert.equal(sheet.name, 'Ledger')
  assert.equal(sheet.rowCount, 166)
  assert.equal(sheet.colCount, 16)
  assert.deepEqual(sheet.freeze, { rows: 1, cols: 1 })
  // Only the cells ExcelJS actually holds, not a rectangle walk: this workbook styles every one of
  // its 166 rows across 16 columns, so all of them are materialised.
  assert.equal(sheet.cells.size, 166 * 16)

  // Width 43 and 14 are what the fixture writes to <col>; Excel's char-to-pixel rule gives 306 and 103.
  assert.equal(sheet.colWidths.get(0), 306)
  assert.equal(sheet.colWidths.get(1), 103)
  assert.equal(sheet.colWidths.get(15), 103)
  // Only the header carries a custom height (30 pt); the body rows use the sheet default.
  assert.equal(sheet.rowHeights.get(0), 40)
  assert.equal(sheet.rowHeights.has(1), false)
})

test('styled.xlsx: A2 fill and thin borders, number formats per column', async () => {
  const [sheet] = (await loadFixture('styled.xlsx')).sheets

  const a2 = sheet.cells.get('1,0')
  assert.equal(a2.value.value, '4002_Store 20000000000000')
  assert.equal(a2.css.fill, '#F2F5F8')
  for (const side of ['top', 'right', 'bottom', 'left']) {
    assert.deepEqual(a2.css.border[side], { style: 'thin', color: '#B0B0B0' })
  }
  assert.equal(a2.css.font.family, 'Calibri')
  assert.equal(a2.css.font.size, 11)

  const b2 = sheet.cells.get('1,1')
  assert.equal(b2.value.kind, 'number')
  assert.equal(b2.value.value, 200274)
  assert.equal(b2.value.display, '200,274.00')

  // The formats this fixture assigns per column: '#,##0', '0.00%' and '#,##0.0'.
  assert.equal(sheet.cells.get('1,3').value.display, '250,200')
  assert.equal(sheet.cells.get('1,5').value.display, '2.00%')
  assert.equal(sheet.cells.get('1,7').value.display, '82.0')
  // E2/E3 are formulas the fixture writes without a cached result, so they display their formula text.
  assert.equal(sheet.cells.get('1,4').value.formula, 'B2-D2')

  const a1 = sheet.cells.get('0,0')
  assert.equal(a1.css.fill, '#1F4E79')
  assert.equal(a1.css.font.color, '#FFFFFF')
  assert.equal(a1.css.font.bold, true)
  // The painted header carries the same thin box as the body, not just the fill.
  assert.deepEqual(a1.css.border.top, { style: 'thin', color: '#B0B0B0' })
})

test('loading the fixtures stays well under 300 ms', async () => {
  for (const name of ['styled.xlsx', 'rich.xlsx']) {
    const bytes = bytesOf(name)
    const runs = []
    let rows = 0
    for (let run = 0; run < 5; run++) {
      const started = performance.now()
      const model = await loadWorkbook(bytes)
      runs.push(performance.now() - started)
      rows = Math.max(rows, ...model.sheets.map((sheet) => sheet.rowCount))
    }
    const best = Math.min(...runs)
    console.log(`${name}: ${rows} rows, ${best.toFixed(1)} ms best of ${runs.length} (${runs.map((ms) => ms.toFixed(1)).join(', ')})`)
    assert.ok(best < 300, `${name} took ${best.toFixed(1)} ms`)
  }
})

test('styled.xlsx: the 166x16 sheet materialises only existing cells', async () => {
  const [sheet] = (await loadFixture('styled.xlsx')).sheets
  // Every one of the 166 rows is styled across all 16 columns in this workbook, so the map holds
  // exactly the cells ExcelJS keeps; nothing is invented beyond the last row or column.
  assert.equal(sheet.cells.size, 166 * 16)
  assert.equal(sheet.cells.has('166,0'), false)
  assert.equal(sheet.cells.has('0,16'), false)
})

/* ------------------------------------------------------------------ numfmt --------------------- */

test('numfmt: the codes working books actually carry', () => {
  const cases = [
    ['general', 42, 'General', false, '42'],
    ['general keeps 4 significant digits of float noise', 0.08119999999999999, 'General', false, '0.0812'],
    ['grouped integer', 282349, '#,##0', false, '282,349'],
    ['grouped with decimals', 282349, '#,##0.00', false, '282,349.00'],
    ['large grouped', 13755522, '#,##0.00', false, '13,755,522.00'],
    ['percent with decimals', 0.0812, '0.00%', false, '8.12%'],
    ['percent integer', 0.0812, '0%', false, '8%'],
    ['percent one decimal', 0.0812, '0.0%', false, '8.1%'],
    ['scientific', 1234567.891, '0.00E+00', false, '1.23E+06'],
    ['currency literal', 1234.5, '#,##0.00 "RUB"', false, '1,234.50 RUB'],
    ['currency via [$...]', 1234.5, '#,##0.00 [$₽-419]', false, '1,234.50 ₽'],
    ['negative with a minus', -1234.5, '#,##0.00', false, '-1,234.50'],
    ['negative in brackets', -1234.5, '#,##0.00;[Red](#,##0.00)', false, '(1,234.50)'],
    ['zero section', 0, '#,##0.00;(#,##0.00);"zero"', false, 'zero'],
    ['zero as a dash', 0, '#,##0.00;(#,##0.00);"-"', false, '-'],
    ['date yyyy-mm-dd', 43831, 'yyyy-mm-dd', true, '2020-01-01'],
    ['date dd.mm.yyyy', 43831, 'dd.mm.yyyy', true, '01.01.2020'],
    ['date d mmm yyyy', 43831, 'd mmm yyyy', true, '1 Jan 2020'],
    ['month name only', 43831, 'mmm', true, 'Jan'],
    ['time hh:mm', 43831.5, 'hh:mm', true, '12:00'],
    ['elapsed hours', 1.5, '[h]:mm:ss', true, '36:00:00'],
    ['date and time with am/pm', 43831.75, 'm/d/yy h:mm AM/PM', true, '1/1/20 6:00 PM'],
    ['text section keeps the text', 'abc', '"n/a";;;@', false, 'abc'],
    ['unknown code degrades to General', 1234, 'zzz', false, '1234'],
    ['empty format is General', 1234, undefined, false, '1234']
  ]

  for (const [label, value, fmt, isDate, expected] of cases) {
    assert.equal(formatValue(value, fmt, isDate), expected, label)
  }
})

test('numfmt: Date values and date-format detection', () => {
  // Excel serial 43831 is 2020-01-01; ExcelJS hands dates over as JS Date objects at UTC midnight.
  assert.equal(formatValue(new Date(Date.UTC(2020, 0, 1)), 'yyyy-mm-dd'), '2020-01-01')
  assert.equal(formatValue(new Date(Date.UTC(2020, 0, 1, 18, 30)), 'dd.mm.yyyy hh:mm'), '01.01.2020 18:30')
  assert.equal(formatValue(null, '#,##0'), '')
  assert.equal(formatValue(true, 'General'), 'TRUE')

  assert.equal(isDateFormat('yyyy-mm-dd'), true)
  assert.equal(isDateFormat('[h]:mm:ss'), true)
  assert.equal(isDateFormat('mmm'), true)
  assert.equal(isDateFormat('General'), false)
  assert.equal(isDateFormat('#,##0.00 "RUB"'), false)
  assert.equal(isDateFormat('0.00%'), false)
})

/* ------------------------------------------------------------------ style ---------------------- */

test('style: theme palette and tint resolve to the ARGB Excel would show', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Theme')
  sheet.getCell('A1').value = 1
  sheet.getCell('A1').font = { color: { theme: 5, tint: 0.4 } }
  sheet.getCell('A2').value = 2
  sheet.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4, tint: -0.25 } }
  sheet.getCell('A3').value = 3
  sheet.getCell('A3').font = { color: { argb: '80FF0000' } }

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const cells = model.sheets[0].cells

  // Theme index 5 is accent2 (#C0504D) in the default Office theme; 40% lighter is #D99694.
  assert.equal(cells.get('0,0').css.font.color, '#D99694')
  // Theme index 4 is accent1 (#4F81BD); 25% darker is the Excel swatch #376092.
  assert.equal(cells.get('1,0').css.fill, '#376092')
  // Alpha is dropped: a cell colour is opaque.
  assert.equal(cells.get('2,0').css.font.color, '#FF0000')
})

test('style: an unresolvable theme colour is reported, not guessed', () => {
  const css = resolveStyle({ font: { color: { theme: 40 } } }, themePalette(undefined))

  assert.equal(css.font.color, null)
  assert.equal(css.notes.length, 1)
  assert.match(css.notes[0], /theme colour 40/)
})

test('style: a style-less cell resolves to an empty, non-null css block', () => {
  const css = resolveStyle(undefined, null)

  assert.equal(css.fill, null)
  assert.deepEqual(css.border, { top: null, right: null, bottom: null, left: null })
  assert.equal(css.font.bold, false)
  assert.equal(css.align.wrapText, false)
})

/* ------------------------------------------------------------------ shapes and edges ----------- */

test('an empty sheet gives an empty model instead of throwing', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.addWorksheet('Empty')

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const [sheet] = model.sheets

  assert.equal(sheet.name, 'Empty')
  assert.equal(sheet.rowCount, 0)
  assert.equal(sheet.colCount, 0)
  assert.equal(sheet.cells.size, 0)
  assert.deepEqual([...sheet.colWidths], [])
  assert.deepEqual([...sheet.rowHeights], [])
  assert.deepEqual(sheet.merges, [])
  assert.deepEqual(sheet.freeze, { rows: 0, cols: 0 })
  assert.deepEqual([...sheet.hidden.rows], [])
  assert.deepEqual([...sheet.hidden.cols], [])
})

test('a sheet without styles still exposes its values', async () => {
  const [, notes] = (await loadFixture('rich.xlsx')).sheets

  assert.equal(notes.cells.size, 2)
  const a1 = notes.cells.get('0,0')
  assert.equal(a1.value.kind, 'text')
  assert.equal(a1.value.display, 'Sheet with no styling')
  assert.equal(a1.css.fill, null)
  assert.equal(a1.css.font.bold, false)
  assert.equal(a1.css.font.color, null)
  assert.deepEqual(a1.css.border, { top: null, right: null, bottom: null, left: null })
})

test('hidden rows and columns are reported once, by index', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Hidden')
  sheet.getCell('A1').value = 'x'
  sheet.getCell('B2').value = 'y'
  sheet.getRow(2).hidden = true
  sheet.getRow(2).height = 45
  sheet.getColumn(2).hidden = true
  sheet.getColumn(2).width = 20

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const [loaded] = model.sheets

  assert.deepEqual([...loaded.hidden.rows], [1])
  assert.deepEqual([...loaded.hidden.cols], [1])
  assert.equal(loaded.rowHeights.get(1), 60)
  assert.equal(loaded.colWidths.get(1), 20 * 7 + 5)
})

test('a styled but empty cell is kept, an unstyled empty one is dropped', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Styled')
  sheet.getCell('A1').value = 'x'
  sheet.getCell('B1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } }
  sheet.getCell('D1').value = null
  sheet.getCell('D1').border = { bottom: { style: 'double', color: { argb: 'FF000000' } } }

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const cells = model.sheets[0].cells

  const blank = cells.get('0,1')
  assert.equal(blank.value.kind, 'blank')
  assert.equal(blank.value.display, '')
  assert.equal(blank.css.fill, '#FFC7CE')
  assert.deepEqual(cells.get('0,3').css.border.bottom, { style: 'double', color: '#000000' })
  // Column C between them holds nothing and stays out of the map.
  assert.equal(cells.has('0,2'), false)
  assert.equal(cells.size, 3)
})

test('shared formulas carry their translated text and cached results', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Shared')
  sheet.getCell('B1').value = 10
  sheet.getCell('A1').value = { formula: 'B1*2', shareType: 'shared', ref: 'A1:A3', result: 20 }
  sheet.getCell('A2').value = { sharedFormula: 'A1', result: 0 }
  sheet.getCell('A3').value = { sharedFormula: 'A1', result: 0 }
  sheet.getCell('C1').value = { error: '#DIV/0!' }
  sheet.getCell('D1').value = { richText: [{ text: 'a', font: { bold: true } }, { text: 'b' }] }

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const cells = model.sheets[0].cells

  assert.equal(cells.get('0,0').value.formula, 'B1*2')
  assert.equal(cells.get('0,0').value.display, '20')
  // The master's formula, translated for the slave cell.
  assert.equal(cells.get('1,0').value.formula, 'B2*2')
  assert.equal(cells.get('2,0').value.formula, 'B3*2')

  assert.equal(cells.get('0,2').value.kind, 'error')
  assert.equal(cells.get('0,2').value.display, '#DIV/0!')

  assert.equal(cells.get('0,3').value.kind, 'text')
  assert.equal(cells.get('0,3').value.value, 'ab')
  assert.equal(cells.get('0,3').value.rich, true)
})

/* ------------------------------------------------------------------ absolute rel targets ------ */

/**
 * An .xlsx whose relationship targets are absolute part names, which OPC allows and Excel writes.
 * ExcelJS indexes the table parts by their literal target, so the sheet's `tableParts` entry cannot
 * be resolved and the whole load throws.
 */
async function absoluteTargetBook() {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Stock')
  sheet.addTable({
    name: 'Inventory',
    ref: 'A1',
    headerRow: true,
    columns: [{ name: 'Item' }, { name: 'Qty' }],
    rows: [['Coffee', 12], ['Tea', 7]]
  })

  const files = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()))
  const patch = (name, from, to) => {
    const text = new TextDecoder().decode(files[name])
    assert.ok(text.includes(from), `${name} carries ${from}`)
    files[name] = new TextEncoder().encode(text.replaceAll(from, to))
  }
  patch('xl/worksheets/_rels/sheet1.xml.rels', 'Target="../tables/table1.xml"', 'Target="/xl/tables/table1.xml"')
  patch('xl/_rels/workbook.xml.rels', 'Target="worksheets/sheet1.xml"', 'Target="/xl/worksheets/sheet1.xml"')
  return zipSync(files)
}

test('a relationship with an absolute Target opens instead of throwing', async () => {
  const bytes = await absoluteTargetBook()
  const untouched = Uint8Array.from(bytes)

  // The bug this guards: `tableParts` resolves to undefined and ExcelJS dies on `table.name`.
  await assert.rejects(new ExcelJS.Workbook().xlsx.load(bytes), /reading 'name'/)

  const model = await loadWorkbook(bytes)
  assert.deepEqual(model.sheets.map((sheet) => sheet.name), ['Stock'])

  const [sheet] = model.sheets
  assert.equal(sheet.rowCount, 3)
  assert.equal(sheet.colCount, 2)
  assert.equal(sheet.cells.get('0,0').value.display, 'Item')
  assert.equal(sheet.cells.get('1,0').value.display, 'Coffee')
  assert.equal(sheet.cells.get('1,1').value.display, '12')
  assert.equal(sheet.cells.get('2,1').value.display, '7')

  // Only the copy handed to ExcelJS is rewritten; the document the caller holds is untouched.
  assert.deepEqual(bytes, untouched)
})

test('only the worksheet relationship part is rewritten, the rest is copied verbatim', async () => {
  const bytes = await absoluteTargetBook()
  const rebuilt = normalizeAbsoluteTargets(bytes)
  assert.ok(rebuilt instanceof Uint8Array, 'the book needs rewriting')

  const before = unzipSync(bytes)
  const after = unzipSync(rebuilt)
  assert.deepEqual(Object.keys(after), Object.keys(before))

  const rels = new TextDecoder().decode(after['xl/worksheets/_rels/sheet1.xml.rels'])
  assert.ok(rels.includes('Target="../tables/table1.xml"'), rels)
  assert.ok(!rels.includes('/xl/tables/'), rels)

  // A workbook-level absolute target is legal and ExcelJS resolves it itself: left untouched.
  const workbookRels = new TextDecoder().decode(after['xl/_rels/workbook.xml.rels'])
  assert.ok(workbookRels.includes('Target="/xl/worksheets/sheet1.xml"'), workbookRels)

  for (const name of Object.keys(before)) {
    if (name === 'xl/worksheets/_rels/sheet1.xml.rels') continue
    assert.deepEqual(after[name], before[name], name)
  }
})

test('a workbook without absolute targets is not rebuilt', async () => {
  const bytes = bytesOf('styled.xlsx')
  assert.equal(normalizeAbsoluteTargets(bytes), null)
  assert.equal(normalizeAbsoluteTargets(new Uint8Array([0x50, 0x4b, 0x03, 0x04])), null, 'not a zip')

  // What the check costs is the inflation of the few `.rels` parts; the archive itself is never
  // parsed and no copy is built, so it has to stay a fraction of one load.
  let probe = Infinity
  for (let run = 0; run < 5; run++) {
    const started = performance.now()
    normalizeAbsoluteTargets(bytes)
    probe = Math.min(probe, performance.now() - started)
  }
  let load = Infinity
  for (let run = 0; run < 5; run++) {
    const started = performance.now()
    await loadWorkbook(bytes)
    load = Math.min(load, performance.now() - started)
  }
  assert.ok(probe < load * 0.2, `${probe.toFixed(2)} ms to check, ${load.toFixed(1)} ms to load`)
})

// A workbook from the 481-file audit: two sheets whose table parts are reached through absolute
// targets. Absent on another machine, so the test reports itself skipped rather than failing.
const AUDIT_WORKBOOK = 'C:/Users/ВкусВилл/Downloads/Корректировки кафе ежедневный отчет/output/Корректировки_кафе_2026-10-22.xlsx'

test('the audit workbook with absolute table targets opens', async (t) => {
  if (!existsSync(AUDIT_WORKBOOK)) return t.skip('the audit workbook is not on this machine')

  const bytes = new Uint8Array(readFileSync(AUDIT_WORKBOOK))
  assert.ok(normalizeAbsoluteTargets(bytes) instanceof Uint8Array, 'needs rewriting')

  const model = await loadWorkbook(bytes)
  assert.deepEqual(model.sheets.map((sheet) => sheet.name), ['500 Минус-корректировки', '510 Инвентаризация'])
  assert.deepEqual(model.sheets.map((sheet) => [sheet.rowCount, sheet.colCount]), [[18, 6], [18, 6]])
  assert.deepEqual(model.sheets.map((sheet) => sheet.cells.size), [82, 82])
  assert.equal(model.sheets[0].cells.get('4,4').value.display, '-3,208.87')
})

/* ------------------------------------------------------------------ group separators --------- */

test('numfmt: a literal separator inside the integer part groups digits', () => {
  // Excel holds the separator at the position the pattern gives it: the digits that do not fit are
  // prepended, so 1234567 under `#\ ##0` is `1234 567`, not `1 234 567`.
  const cases = [
    ['escaped space', 258429.9, '#\\ ##0', '258 430'],
    ['escaped space with a literal suffix', 258429.9, '#\\ ##0"   "', '258 430   '],
    ['four digits', 1234, '#\\ ##0', '1 234'],
    ['seven digits', 1234567, '#\\ ##0', '1234 567'],
    ['negative', -1234, '#\\ ##0', '-1 234'],
    ['zero keeps its separator', 0, '#\\ ##0', ' 0'],
    ['three groups', 123456789, '[$-409]###\\ ###\\ ###', '123 456 789'],
    ['three groups, long', 1234567890, '###\\ ###\\ ###', '1234 567 890'],
    ['no-break space', 1234567, '#\u00a0##0', '1234\u00a0567'],
    ['apostrophe', 1234567, "#\\'##0", "1234'567"],
    ['decimals after a space group', 258429.9, '#\\ ##0.00', '258 429.90'],
    // A literal that is not a separator stays out of the number, as it always did.
    ['letter dropped', 1234, '#x##0', '1234'],
    ['separator kept, letter dropped', 1234, '#\\ x##0', '1 234']
  ]

  for (const [label, value, fmt, expected] of cases) {
    assert.equal(formatValue(value, fmt), expected, label)
  }
})

/* ------------------------------------------------------------------ date sections ------------ */

test('numfmt: a date format is split into sections before it is tokenised', () => {
  // `;@` is the text section, not a literal tail: Excel shows `1 Jan`, not `1 Jan;@`.
  assert.equal(formatValue(43831, '[$-419]d\\ mmm;@', true), '1 Jan')
  assert.equal(formatValue(43831, 'dd/mm/yy;@', true), '01/01/20')
  assert.equal(formatValue(43831, 'yyyy-mm-dd', true), '2020-01-01')
  assert.equal(formatValue(43831, '[$-409]mmm\\-yy;@', true), 'Jan-20')
  // The same format still renders the text section for a text value.
  assert.equal(formatValue('abc', 'd mmm;@'), 'abc')

  assert.equal(isDateFormat('[$-419]d\\ mmm;@'), true)
  assert.equal(isDateFormat('0.00" d"'), false)
})

/* ------------------------------------------------------------------ conditions and scaling --- */

test('numfmt: conditional sections pick the first section whose test passes', () => {
  const scale = '[>=1000000]#,##0.0,," млн";[>1000]#,###.00," тыс.";0'
  assert.equal(formatValue(1, scale), '1')
  assert.equal(formatValue(999, scale), '999')
  assert.equal(formatValue(5000, scale), '5.00 тыс.')
  assert.equal(formatValue(1234567, scale), '1.2 млн')

  // Without a condition the sign rule still decides.
  assert.equal(formatValue(3, '[>3]0.0;[<=3]0.00'), '3.00')
  assert.equal(formatValue(-5, '[<0]"neg";0'), 'neg')
  // No section matches and none is unconditional: Excel falls back to General.
  assert.equal(formatValue(1, '[>3]0'), '1')
})

test('numfmt: trailing commas scale the number and never become decimals', () => {
  assert.equal(formatValue(1234567, '#,##0,,', false), '1')
  assert.equal(formatValue(1234567, '#,##0,"K"'), '1,235K')
  assert.equal(formatValue(1500, '#,##0,'), '2')
  // The same format without a condition, to keep the scaling rule independent of it.
  assert.equal(formatValue(1250000, '#,##0.0,'), '1,250.0')
})

/* ------------------------------------------------------------------ literal zero section ------ */

test('numfmt: a zero section made of a literal prints as written', () => {
  assert.equal(formatValue(0, '#,##0;[Red]\\-#,##0;\\–'), '–')
  assert.equal(formatValue(0, '#,##0;\\(#,##0\\);\\-'), '-')
  assert.equal(formatValue(-1234, '#,##0;[Red]\\-#,##0;\\–'), '-1,234')
  // Quoted literals always rendered; plain characters are still an unknown code, not a literal.
  assert.equal(formatValue(0, '#,##0.00;(#,##0.00);"zero"'), 'zero')
  assert.equal(formatValue(1234, 'zzz'), '1234')
})

/* ------------------------------------------------------------------ `?` placeholders --------- */

test('numfmt: ? reserves the width of a digit it did not print', () => {
  const accounting = '_-* #,##0.00\\ _₽_-;\\-* #,##0.00\\ _₽_-;_-* "-"??\\ _₽_-;_-@_-'
  // Excel 16 renders this zero as ` -` plus five spaces: two from `??`, three from `\ `, `_₽`, `_-`.
  assert.equal(formatValue(0, accounting), ' -     ')
  assert.equal(formatValue(1234.5, accounting), ' 1,234.50   ')
  assert.equal(formatValue(-1234.5, accounting), '-1,234.50   ')

  assert.equal(formatValue(0, '??'), '  ')
  assert.equal(formatValue(1.5, '??0.00'), '  1.50')
  assert.equal(formatValue(1.5, '0.0?'), '1.5 ')
  assert.equal(formatValue(0, '#'), '')
})

/* ------------------------------------------------------------------ 1904 date system --------- */

test('numfmt: the 1904 epoch moves the serials by 1462 days', () => {
  // Excel 16, a 1904 workbook: serial 43831 is 2024-01-02 and 0 is 1904-01-01.
  assert.equal(formatValue(43831, 'yyyy-mm-dd', true, true), '2024-01-02')
  assert.equal(formatValue(0, 'yyyy-mm-dd', true, true), '1904-01-01')
  assert.equal(formatValue(1462, 'yyyy-mm-dd', true, true), '1908-01-02')
  // A Date is turned into a serial of the same system and comes back unchanged.
  assert.equal(formatValue(new Date(Date.UTC(2024, 0, 15)), 'yyyy-mm-dd', true, true), '2024-01-15')
  // Elapsed time does not depend on the epoch.
  assert.equal(formatValue(1.5, '[h]:mm:ss', true, true), '36:00:00')
  // The default stays the 1900 system.
  assert.equal(formatValue(43831, 'yyyy-mm-dd', true), '2020-01-01')
})

test('date1904: a 1904 workbook formats its serials in the same system', async () => {
  const workbook = new ExcelJS.Workbook()
  workbook.properties.date1904 = true
  const sheet = workbook.addWorksheet('Dated')
  sheet.getCell('A1').value = 100
  sheet.getCell('A1').numFmt = 'yyyy-mm-dd'
  sheet.getCell('A2').value = new Date(Date.UTC(2024, 0, 15))
  sheet.getCell('A2').numFmt = 'yyyy-mm-dd'
  sheet.getCell('A3').value = 1.5
  sheet.getCell('A3').numFmt = '[h]:mm:ss'

  const model = await loadWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()))
  assert.equal(model.date1904, true)

  const [dated] = model.sheets
  assert.equal(dated.cells.get('0,0').value.display, '1904-04-10')
  assert.equal(dated.cells.get('1,0').value.display, '2024-01-15')
  assert.equal(dated.cells.get('2,0').value.display, '36:00:00')
})
