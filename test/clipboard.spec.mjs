// Clipboard formatting: the pure formatters plus copyText against a stub document.
// The fixture case loads a real workbook through the model (hence `exceljs`, hence `.spec`).
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { selectionToTsv, selectionToCsv, selectionToMarkdown, copyText } from '../src/ui/clipboard.js'
import { loadWorkbook } from '../src/model/workbook.js'
import { fixtureBytes } from './helpers/fixtures.mjs'

const text = (display, formula = null) => ({ kind: 'text', value: display, display, formula })
const number = (value, display, formula = null) => ({ kind: 'number', value, display, formula })
const blank = () => ({ kind: 'blank', value: null, display: '', formula: null })

/** A SheetModel-shaped fixture: `rows` is a grid of CellValue or null for an absent cell. */
function sheetOf(rows, name = 'Sheet1') {
  const cells = new Map()
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      if (!value) return
      cells.set(`${r},${c}`, {
        ref: `${r},${c}`,
        row: r,
        col: c,
        value,
        css: { font: {}, fill: null, border: {}, align: {} }
      })
    })
  })
  return {
    name,
    index: 0,
    rowCount: rows.length,
    colCount: rows.reduce((max, row) => Math.max(max, row.length), 0),
    cells,
    colWidths: new Map(),
    rowHeights: new Map(),
    merges: [],
    freeze: { rows: 0, cols: 0 },
    hidden: { rows: new Set(), cols: new Set() }
  }
}

/** RFC 4180 reader, small enough to trust: used to prove the writer round-trips. */
function parseCsv(text_) {
  const rows = []
  let row = []
  let field = ''
  let quoted = false
  for (let i = 0; i < text_.length; i++) {
    const char = text_[i]
    if (quoted) {
      if (char === '"' && text_[i + 1] === '"') {
        field += '"'
        i++
      } else if (char === '"') {
        quoted = false
      } else {
        field += char
      }
    } else if (char === '"') {
      quoted = true
    } else if (char === ',') {
      row.push(field)
      field = ''
    } else if (char === '\r' && text_[i + 1] === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
      i++
    } else {
      field += char
    }
  }
  row.push(field)
  rows.push(row)
  return rows
}

test('tsv: rows by newline, cells by tab, absent cells are empty fields', () => {
  const sheet = sheetOf([
    [text('a'), null],
    [null, text('b')],
    [text('c'), text('d')]
  ])
  assert.equal(selectionToTsv(sheet, { r1: 0, c1: 0, r2: 2, c2: 1 }), 'a\t\n\tb\nc\td')
})

test('tsv: a tab or line break in a value becomes a space so one cell stays one field', () => {
  const sheet = sheetOf([[text('one\ttwo\nthree')], [text('line\r\nbreak')]])
  assert.equal(selectionToTsv(sheet, { r1: 0, c1: 0, r2: 1, c2: 0 }), 'one two three\nline break')
})

test('tsv: the formatted display is what gets copied, not the raw number', () => {
  const sheet = sheetOf([[number(282349, '282,349.00'), blank()]])
  assert.equal(selectionToTsv(sheet, { r1: 0, c1: 0, r2: 0, c2: 1 }), '282,349.00\t')
})

test('csv: commas and quotes are escaped per RFC 4180 and the text round-trips', () => {
  const sheet = sheetOf([
    [text('a,b'), text('say "hi"')],
    [text('plain'), text('line\nbreak')]
  ])
  const selection = { r1: 0, c1: 0, r2: 1, c2: 1 }
  const csv = selectionToCsv(sheet, selection)
  assert.equal(csv, '"a,b","say ""hi"""\r\nplain,"line\nbreak"')
  assert.deepEqual(parseCsv(csv), [
    ['a,b', 'say "hi"'],
    ['plain', 'line\nbreak']
  ])
})

test('csv: plain values stay unquoted and records are separated by CRLF', () => {
  const sheet = sheetOf([[text('x'), text('y')], [text('1'), text('2')]])
  assert.equal(selectionToCsv(sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }), 'x,y\r\n1,2')
})

test('markdown: header, separator and data, numeric column right-aligned', () => {
  const sheet = sheetOf([
    [text('Магазин'), text('Продажи')],
    [text('A'), number(282349, '282,349.00')]
  ])
  assert.equal(
    selectionToMarkdown(sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }),
    ['| Магазин | Продажи |', '| --- | ---: |', '| A | 282,349.00 |'].join('\n')
  )
})

test('markdown: pipes are escaped and line breaks become <br>', () => {
  const sheet = sheetOf([
    [text('h1'), text('h2')],
    [text('a|b'), text('x\ny')]
  ])
  assert.equal(
    selectionToMarkdown(sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }),
    ['| h1 | h2 |', '| --- | --- |', '| a\\|b | x<br>y |'].join('\n')
  )
})

test('markdown: a single row is written as the header GFM requires', () => {
  const sheet = sheetOf([[text('only'), number(1, '1.00')]])
  assert.equal(selectionToMarkdown(sheet, { r1: 0, c1: 0, r2: 0, c2: 1 }), '| only | 1.00 |\n| --- | ---: |')
})

test('formulas: a formula cell copies as =FORMULA only when asked', () => {
  const sheet = sheetOf([[number(15, '15', 'SUM(B2:B6)'), text('10')]])
  const selection = { r1: 0, c1: 0, r2: 0, c2: 1 }
  assert.equal(selectionToTsv(sheet, selection, { formulas: true }), '=SUM(B2:B6)\t10')
  assert.equal(selectionToTsv(sheet, selection), '15\t10')
  assert.equal(selectionToCsv(sheet, selection, { formulas: true }), '=SUM(B2:B6),10')
  assert.equal(selectionToMarkdown(sheet, selection, { formulas: true }), '| =SUM(B2:B6) | 10 |\n| --- | --- |')
})

test('empty selection is an empty string for every format', () => {
  const sheet = sheetOf([[text('a')]])
  const empty = { r1: 0, c1: 0, r2: -1, c2: -1 }
  for (const format of [selectionToTsv, selectionToCsv, selectionToMarkdown]) {
    assert.equal(format(sheet, null), '')
    assert.equal(format(sheet, empty), '')
    assert.equal(format(null, { r1: 0, c1: 0, r2: 0, c2: 0 }), '')
  }
})

/** Minimal document for copyText: enough DOM to show the textarea is filled, used and removed. */
function stubDocument(copyResult) {
  const body = { children: [] }
  const doc = {
    activeElement: null,
    body,
    selected: null,
    command: null,
    copyData: null,
    prevented: false,
    execCommand(command) {
      doc.command = command
      const node = doc.selected
      if (node?.handlers.copy) {
        node.handlers.copy({
          clipboardData: { setData: (type, value) => { doc.copyData = [type, value] } },
          preventDefault: () => { doc.prevented = true }
        })
      }
      return copyResult
    },
    createElement(tagName) {
      const node = {
        tagName,
        value: '',
        style: {},
        attributes: {},
        handlers: {},
        setAttribute(name, value) {
          node.attributes[name] = String(value)
        },
        addEventListener(type, fn) {
          node.handlers[type] = fn
        },
        removeEventListener(type) {
          delete node.handlers[type]
        },
        select() {
          doc.selected = node
        },
        remove() {
          const index = body.children.indexOf(node)
          if (index >= 0) body.children.splice(index, 1)
        }
      }
      return node
    }
  }
  body.append = (node) => body.children.push(node)
  return doc
}

test('copyText: fills a scratch textarea, runs the legacy copy command and cleans up', () => {
  const doc = stubDocument(true)
  assert.equal(copyText('a\tb', doc), true)
  assert.equal(doc.command, 'copy')
  assert.deepEqual(doc.copyData, ['text/plain', 'a\tb'])
  assert.equal(doc.prevented, true)
  assert.deepEqual(doc.body.children, [])
})

test('copyText: reports failure when the frame refuses the command', () => {
  assert.equal(copyText('x', stubDocument(false)), false)
  assert.equal(copyText('x', {}), false)
})

test('styled workbook: Ledger A1:C3 copies as the grid shows it', async () => {
  const bytes = fixtureBytes('styled.xlsx')
  const model = await loadWorkbook(bytes)
  const sheet = model.sheets[0]
  const selection = { r1: 0, c1: 0, r2: 2, c2: 2 }

  // Column C holds formulas the fixture writes without a cached result, so the grid (and therefore
  // the clipboard) shows the formula text there; columns A and B show their formatted values.
  assert.equal(
    selectionToTsv(sheet, selection),
    [
      'Store\tSales\tRevenue',
      '4002_Store 20000000000000\t200,274.00\tB2*48.7',
      '4003_Store 30000000000000\t200,411.00\tB3*48.7'
    ].join('\n')
  )
  assert.equal(
    selectionToMarkdown(sheet, selection),
    [
      '| Store | Sales | Revenue |',
      '| --- | ---: | --- |',
      '| 4002_Store 20000000000000 | 200,274.00 | B2*48.7 |',
      '| 4003_Store 30000000000000 | 200,411.00 | B3*48.7 |'
    ].join('\n')
  )
  const csv = parseCsv(selectionToCsv(sheet, selection))
  assert.equal(csv.length, 3)
  assert.deepEqual(csv[1], ['4002_Store 20000000000000', '200,274.00', 'B2*48.7'])
})
