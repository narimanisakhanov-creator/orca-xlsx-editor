// Clipboard formatting: the pure formatters plus copyText against a stub document.
// The fixture case loads a real workbook through the model (hence `exceljs`, hence `.spec`).
import { test } from 'node:test'
import assert from 'node:assert/strict'
import {
  selectionToTsv,
  selectionToCsv,
  selectionToMarkdown,
  selectionToHtml,
  selectionFitsTextBudget,
  copyText,
  copyHtml,
  MAX_HTML_CELLS,
  MAX_HTML_BYTES,
  MAX_HTML_COLSPAN,
  MAX_HTML_ROWSPAN
} from '../src/ui/clipboard.js'
import { loadWorkbook } from '../src/model/workbook.js'
import { fixtureBytes } from './helpers/fixtures.mjs'

const text = (display, formula = null) => ({ kind: 'text', value: display, display, formula })
const number = (value, display, formula = null) => ({ kind: 'number', value, display, formula })
const blank = () => ({ kind: 'blank', value: null, display: '', formula: null })

/** A SheetModel-shaped fixture: `rows` is a grid of CellValue or null for an absent cell. */
function sheetOf(rows, name = 'Sheet1') {
  const cells = new Map()
  rows.forEach((row, r) => {
    row.forEach((entry, c) => {
      if (!entry) return
      // A plain CellValue is the common case; `{ value, css }` lets a test attach resolved styles.
      const value = entry.css ? entry.value : entry
      cells.set(`${r},${c}`, {
        ref: `${r},${c}`,
        row: r,
        col: c,
        value,
        css: entry.css ?? { font: {}, fill: null, border: {}, align: {} }
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

test('text budget refuses repeated long shared values before formatting and accepts small selections', () => {
  const shared = text('x'.repeat(1000))
  const sheet = sheetOf([[shared, shared], [shared, shared]])
  const box = { r1: 0, c1: 0, r2: 1, c2: 1 }
  assert.equal(selectionFitsTextBudget(sheet, box, { maxChars: 400 }), false)
  assert.equal(selectionFitsTextBudget(sheet, box, { maxChars: 20000 }), true)
  assert.equal(selectionFitsTextBudget(sheet, box, { maxChars: 0 }), false)
  assert.equal(selectionFitsTextBudget(sheetOf([[text('x', 'y'.repeat(1000))]]), { r1: 0, c1: 0, r2: 0, c2: 0 }, { formulas: true, maxChars: 400 }), false)
})

test('large repeated shared strings are refused without traversing the whole selection', () => {
  const shared = text('x'.repeat(32767))
  let reads = 0
  const sheet = { cells: { get() { reads++; return { value: shared } } } }
  const box = { r1: 0, c1: 0, r2: 199999, c2: 0 }
  assert.equal(selectionFitsTextBudget(sheet, box, { maxChars: 8 * 1024 * 1024 }), false)
  assert.ok(reads < 100)
})

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

const TABLE_OPEN = '<table border="0" cellspacing="0" cellpadding="0" style="border-collapse:collapse"><tbody>'

test('html: escaped cells, numbers right-aligned, and an empty cell kept in place', () => {
  const sheet = sheetOf([
    [text('a & b'), number(282349, '282,349.00')],
    [text('<script>alert("x")</script>'), text('line\nbreak')]
  ])
  const html = selectionToHtml(sheet, { r1: 0, c1: 0, r2: 1, c2: 1 })
  assert.equal(
    html,
    [
      TABLE_OPEN,
      '<tr><td>a &amp; b</td><td style="text-align:right">282,349.00</td></tr>',
      '<tr><td>&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;</td><td>line<br>break</td></tr>',
      '</tbody></table>'
    ].join('')
  )
})

test('html: Unicode survives unchanged and only markup characters are escaped', () => {
  const sheet = sheetOf([[text('Магазин №5 — «Ко» & сын 😀')]])
  const html = selectionToHtml(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 })
  assert.match(html, /Магазин №5 — «Ко» &amp; сын 😀/)
  assert.equal((html.match(/<td/g) ?? []).length, 1)
})

test('html: text that looks like a formula is defused, a real one stays a formula when asked', () => {
  const sheet = sheetOf([[text('=1+1'), number(15, '15', 'SUM(B2:B6)')]])
  const selection = { r1: 0, c1: 0, r2: 0, c2: 1 }
  assert.match(selectionToHtml(sheet, selection), /&#39;=1\+1<\/td><td style="text-align:right">15<\/td>/)
  assert.match(
    selectionToHtml(sheet, selection, { formulas: true }),
    /&#39;=1\+1<\/td><td style="text-align:right">=SUM\(B2:B6\)<\/td>/
  )
})

test('html: resolved font, fill, alignment and border become inline CSS', () => {
  const css = {
    font: { bold: true, italic: true, underline: 'double', strike: true, size: 12, color: '#FF0000', family: 'Arial' },
    fill: '#FFFF00',
    border: { top: { style: 'thin', color: '#000000' }, bottom: { style: 'medium', color: '#FF0000' } },
    align: { horizontal: 'center', vertical: 'bottom', wrapText: true, indent: 2 }
  }
  const sheet = sheetOf([[{ value: text('styled'), css }]])
  const html = selectionToHtml(sheet, { r1: 0, c1: 0, r2: 0, c2: 0 })
  for (const declaration of [
    'font-weight:700',
    'font-style:italic',
    'font-size:12pt',
    'color:#FF0000',
    'font-family:&#39;Arial&#39;',
    'text-decoration:underline line-through',
    'text-decoration-style:double',
    'background-color:#FFFF00',
    'text-align:center',
    'vertical-align:bottom',
    'padding-left:22px',
    'white-space:normal',
    'overflow-wrap:anywhere',
    'border-top:1px solid #000000',
    'border-bottom:2px solid #FF0000'
  ]) {
    assert.ok(html.includes(declaration), `missing ${declaration} in ${html}`)
  }
})

test('html: merged cells emit one span cell and omit the covered cells', () => {
  const sheet = sheetOf([
    [text('Title'), null, text('C')],
    [null, null, text('D')],
    [text('x'), text('y'), text('z')]
  ])
  sheet.merges = [{ r1: 0, c1: 0, r2: 1, c2: 1 }]
  assert.equal(
    selectionToHtml(sheet, { r1: 0, c1: 0, r2: 2, c2: 2 }),
    [
      TABLE_OPEN,
      '<tr><td colspan="2" rowspan="2">Title</td><td>C</td></tr>',
      '<tr><td>D</td></tr>',
      '<tr><td>x</td><td>y</td><td>z</td></tr>',
      '</tbody></table>'
    ].join('')
  )
})

test('html: a merge clipped by the selection keeps its value at the visible corner', () => {
  const sheet = sheetOf([[text('Title'), null], [null, null], [text('x'), text('y')]])
  sheet.merges = [{ r1: 0, c1: 0, r2: 1, c2: 1 }]
  assert.equal(
    selectionToHtml(sheet, { r1: 1, c1: 1, r2: 2, c2: 2 }),
    [
      TABLE_OPEN,
      '<tr><td>Title</td><td></td></tr>',
      '<tr><td>y</td><td></td></tr>',
      '</tbody></table>'
    ].join('')
  )
})

test('html: a large selection is one cell per position, not deduplicated or dropped', () => {
  const rows = []
  for (let r = 0; r < 400; r++) {
    rows.push(Array.from({ length: 50 }, (_, c) => text(`${r}:${c}`)))
  }
  const sheet = sheetOf(rows)
  const html = selectionToHtml(sheet, { r1: 0, c1: 0, r2: 399, c2: 49 })
  assert.equal((html.match(/<tr>/g) ?? []).length, 400)
  assert.equal((html.match(/<td/g) ?? []).length, 20000)
  assert.ok(html.includes('399:49'))
})

test('html: a selection above the cell ceiling is refused instead of clipped', () => {
  assert.equal(MAX_HTML_CELLS, 250000)
  const sheet = sheetOf([[text('x')]])
  assert.throws(
    () => selectionToHtml(sheet, { r1: 0, c1: 0, r2: 500, c2: 500 }),
    (error) => error instanceof RangeError && /too large for HTML clipboard/.test(error.message)
  )
})

test('html: serialized bytes are bounded before the table is built', () => {
  assert.equal(MAX_HTML_BYTES, 8 * 1024 * 1024)
  const selection = { r1: 0, c1: 0, r2: 0, c2: 0 }
  const long = sheetOf([[text('x'.repeat(5000))]])
  assert.throws(
    () => selectionToHtml(long, selection, { maxBytes: 512 }),
    (error) => error instanceof RangeError && /bytes/.test(error.message)
  )
  assert.ok(selectionToHtml(long, selection).includes('xxxx'))

  // The limit counts UTF-8 bytes, not UTF-16 code units: 1,000 Cyrillic characters are 2,000 bytes.
  const cyrillic = sheetOf([[text('Ж'.repeat(1000))]])
  assert.throws(
    () => selectionToHtml(cyrillic, selection, { maxBytes: 1500 }),
    (error) => error instanceof RangeError && /bytes/.test(error.message)
  )
})

test('html: merged spans beyond the browser ceilings are refused, not clamped', () => {
  const wide = sheetOf([[text('wide')]])
  wide.merges = [{ r1: 0, c1: 0, r2: 0, c2: MAX_HTML_COLSPAN }]
  assert.throws(
    () => selectionToHtml(wide, { r1: 0, c1: 0, r2: 0, c2: MAX_HTML_COLSPAN + 1 }),
    (error) => error instanceof RangeError && /merged region/.test(error.message)
  )

  const tall = sheetOf([[text('tall')]])
  tall.merges = [{ r1: 0, c1: 0, r2: MAX_HTML_ROWSPAN, c2: 0 }]
  assert.throws(
    () => selectionToHtml(tall, { r1: 0, c1: 0, r2: MAX_HTML_ROWSPAN + 1, c2: 0 }),
    (error) => error instanceof RangeError && /merged region/.test(error.message)
  )
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
    copiedData: [],
    prevented: false,
    execCommand(command) {
      doc.command = command
      const node = doc.selected
      if (node?.handlers.copy) {
        node.handlers.copy({
          clipboardData: {
            setData: (type, value) => {
              doc.copyData = [type, value]
              doc.copiedData.push([type, value])
            }
          },
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

test('copyHtml: puts text/html and the TSV fallback on the same clipboard event', () => {
  const doc = stubDocument(true)
  assert.equal(copyHtml('<table><tr><td>x</td></tr></table>', 'x', doc), true)
  assert.deepEqual(doc.copiedData, [
    ['text/plain', 'x'],
    ['text/html', '<table><tr><td>x</td></tr></table>']
  ])
  assert.equal(doc.prevented, true)
  assert.deepEqual(doc.body.children, [])
  assert.equal(copyHtml('<b>x</b>', 'x', stubDocument(false)), false)
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
