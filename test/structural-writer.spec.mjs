import { test } from 'node:test'
import assert from 'node:assert/strict'
import { applyStructural } from '../src/write/structural.js'

const enc = new TextEncoder()
const dec = new TextDecoder()
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

function workbook() {
  const files = {
    'xl/workbook.xml': `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="Total">Data!$A$2:$A$4</definedName></definedNames></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${ns}"><dimension ref="A1:B4"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2" ht="22" customHeight="1"><c r="A2"><v>2</v></c><c r="B2"><f>A2</f><v>2</v></c></row><row r="4"><c r="A4"><v>4</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A3:B4"/></mergeCells></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><f>Data!A2</f><v>2</v></c></row></sheetData></worksheet>`,
  }
  return Object.entries(files).map(([name, text]) => ({ name, data: enc.encode(text), changed: false }))
}
const text = (entries, name) => dec.decode(entries.find((entry) => entry.name === name).data)
const change = (entries, name, fn) => entries.map((entry) => entry.name === name ? { ...entry, data: enc.encode(fn(dec.decode(entry.data))) } : entry)

const insert = { sheet: 'Data', axis: 'row', action: 'insert', index: 2, count: 1 }

test('insertion shifts cells, merged ranges, names and cross-sheet formulas atomically', () => {
  const original = workbook()
  const result = applyStructural(original, insert)
  assert.equal(result.status, 'ok', result.reason)
  const sheet = text(result.entries, 'xl/worksheets/sheet1.xml')
  assert.match(sheet, /<row r="3" ht="22"/)
  assert.match(sheet, /<c r="A3"><v>2<\/v>/)
  assert.match(sheet, /<f>A3<\/f>/)
  assert.match(sheet, /mergeCell ref="A4:B5"/)
  assert.match(text(result.entries, 'xl/worksheets/sheet2.xml'), /<f>Data!A3<\/f>/)
  assert.match(text(result.entries, 'xl/workbook.xml'), /Data!\$A\$3:\$A\$5/)
  assert.match(text(original, 'xl/worksheets/sheet1.xml'), /<c r="A2">/)
})

test('delete removes selected cells and contracts a merge', () => {
  const result = applyStructural(workbook(), { ...insert, action: 'delete', index: 3 })
  assert.equal(result.status, 'ok', result.reason)
  const sheet = text(result.entries, 'xl/worksheets/sheet1.xml')
  assert.match(sheet, /mergeCell ref="A3:B3"/)
  assert.match(sheet, /<c r="A3"><v>4<\/v>/)
})

test('refuses protected and unsupported structures without mutating package entries', () => {
  for (const extra of ['<sheetProtection/>', '<extLst><ext uri="unknown"/></extLst>']) {
    const entries = change(workbook(), 'xl/worksheets/sheet1.xml', (xml) => xml.replace('</worksheet>', `${extra}</worksheet>`))
    const before = entries.map((entry) => entry.data.slice())
    const result = applyStructural(entries, insert)
    assert.equal(result.status, 'refused')
    assert.deepEqual(entries.map((entry) => entry.data), before)
  }
})

test('column insertion shifts cell references and expands a crossed merge', () => {
  const result = applyStructural(workbook(), { sheet: 'Data', axis: 'column', action: 'insert', index: 2, count: 1 })
  assert.equal(result.status, 'ok', result.reason)
  const sheet = text(result.entries, 'xl/worksheets/sheet1.xml')
  assert.match(sheet, /<c r="C2"><f>A2<\/f>/)
  assert.match(sheet, /mergeCell ref="A3:C4"/)
})

test('out-of-bounds moves and uncertain formula types fail without mutation', () => {
  const cases = [
    change(workbook(), 'xl/worksheets/sheet1.xml', (xml) => xml.replace('r="A4"', 'r="XFD4"')),
    change(workbook(), 'xl/worksheets/sheet2.xml', (xml) => xml.replace('<f>', '<f t="shared" si="0">')),
  ]
  for (const entries of cases) {
    const original = entries.map((entry) => entry.data.slice())
    const result = applyStructural(entries, { sheet: 'Data', axis: 'column', action: 'insert', index: 2, count: 1 })
    assert.equal(result.status, 'refused')
    assert.deepEqual(entries.map((entry) => entry.data), original)
  }
})

test('unrecognized package parts and external workbook relationships refuse structural edits', () => {
  const entries = workbook()
  const unknown = [...entries, { name: 'xl/unknown/references.xml', data: enc.encode('<references/>') }]
  assert.equal(applyStructural(unknown, insert).status, 'refused')
  const external = change(entries, 'xl/_rels/workbook.xml.rels', (xml) =>
    xml.replace('</Relationships>', `<Relationship Id="bad" Type="${rel}/externalLink" Target="outside.xlsx" TargetMode="External"/></Relationships>`))
  assert.equal(applyStructural(external, insert).status, 'refused')
})

test('row insertion updates selections while preserving column formatting', () => {
  const entries = change(workbook(), 'xl/worksheets/sheet1.xml', (xml) =>
    xml.replace('<sheetData>', '<sheetViews><sheetView workbookViewId="0"><selection activeCell="A2" sqref="A2:B4"/></sheetView></sheetViews><cols><col min="1" max="2" width="20" customWidth="1"/></cols><sheetData>'))
  const result = applyStructural(entries, insert)
  assert.equal(result.status, 'ok', result.reason)
  const sheet = text(result.entries, 'xl/worksheets/sheet1.xml')
  assert.match(sheet, /activeCell="A3" sqref="A3:B5"/)
  assert.match(sheet, /<col min="1" max="2" width="20" customWidth="1"\/>/)
})

test('an opaque macro project stays byte-identical through a structural edit', () => {
  const withMacro = change(workbook(), 'xl/_rels/workbook.xml.rels', (xml) =>
    xml.replace('</Relationships>', `<Relationship Id="rIdVba" Type="${rel}/vbaProject" Target="vbaProject.bin"/></Relationships>`))
  const vba = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4])
  const entries = [...withMacro, { name: 'xl/vbaProject.bin', data: vba, changed: false }]
  const result = applyStructural(entries, insert)
  assert.equal(result.status, 'ok', result.reason)
  assert.deepEqual(result.entries.find((entry) => entry.name === 'xl/vbaProject.bin').data, vba)
  assert.match(text(result.entries, 'xl/worksheets/sheet1.xml'), /<c r="A3"><v>2<\/v>/)
})

test('a stale calculation chain is dropped with its relationship and content type', () => {
  const withParts = change(workbook(), 'xl/_rels/workbook.xml.rels', (xml) =>
    xml.replace('</Relationships>', `<Relationship Id="rIdCalc" Type="${rel}/calcChain" Target="calcChain.xml"/></Relationships>`))
  const entries = [
    ...withParts,
    { name: 'xl/calcChain.xml', data: enc.encode(`<calcChain xmlns="${ns}"><c r="A2" i="1"/></calcChain>`), changed: false },
    { name: '[Content_Types].xml', data: enc.encode(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`), changed: false }
  ]
  const result = applyStructural(entries, insert)
  assert.equal(result.status, 'ok', result.reason)
  assert.equal(result.entries.some((entry) => entry.name === 'xl/calcChain.xml'), false)
  assert.doesNotMatch(text(result.entries, 'xl/_rels/workbook.xml.rels'), /calcChain/)
  assert.doesNotMatch(text(result.entries, '[Content_Types].xml'), /calcChain/)
  assert.match(text(result.entries, 'xl/workbook.xml'), /fullCalcOnLoad="1"/)
})
