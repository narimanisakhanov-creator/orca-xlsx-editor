// Slice C tests: surgical xlsx cell edits.
//
// Deliberately independent of src/write/zip.js: fixtures go straight through
// fflate and the entries are hand-built in the ZipEntry shape of the contract.
// The books themselves come from test/helpers/build-fixtures.mjs, which writes OOXML by hand.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { unzipSync, zipSync } from 'fflate';
import ExcelJS from 'exceljs';

import {
  escapeXml,
  findElement,
  findElements,
  getAttr,
  setAttr,
  topLevelNames,
  unescapeXml,
} from '../src/write/xml.js';
import { applyEdits, xlsxPackage } from '../src/write/xlsx-writer.js';
import { loadWorkbook } from '../src/model/workbook.js';
import { fixtureBytes as loadFixture } from './helpers/fixtures.mjs';

const DECODER = new TextDecoder();
const ENCODER = new TextEncoder();

/** Entries in the contract's ZipEntry shape, without touching the zip slice. */
function entriesFrom(bytes) {
  const files = unzipSync(bytes);
  return Object.keys(files)
    .sort()
    .map((name) => ({
      name,
      data: files[name],
      method: 8,
      crc: 0,
      raw: null,
      time: 0,
      date: 0,
      flags: 0,
      changed: false,
    }));
}

function byName(entries) {
  const map = new Map();
  for (const entry of entries) map.set(entry.name, entry);
  return map;
}

function textOf(entry) {
  return DECODER.decode(entry.data);
}

/** Unpack what applyEdits returned, so the assertions read real archive bytes. */
function unpack(entries) {
  const files = {};
  for (const entry of entries) files[entry.name] = entry.data;
  return unzipSync(zipSync(files, { level: 6 }));
}

function sheetPart(entries, sheetPartName) {
  return textOf(byName(entries).get(sheetPartName));
}

function cellXml(xml, ref) {
  for (const el of findElements(xml, 'c')) {
    if (getAttr(el.attrs, 'r') === ref) return xml.slice(el.start, el.end);
  }
  return null;
}

function changedNames(before, after) {
  const a = byName(before);
  const b = byName(after);
  const names = [];
  for (const name of a.keys()) if (!b.has(name)) names.push(`-${name}`);
  for (const name of b.keys()) if (!a.has(name)) names.push(`+${name}`);
  for (const [name, entry] of b) {
    const old = a.get(name);
    if (old && Buffer.compare(Buffer.from(old.data), Buffer.from(entry.data)) !== 0) names.push(name);
  }
  return names.sort();
}

const RICH_SUMMARY = 'xl/worksheets/sheet1.xml';

test('xml helpers keep offsets, attrs and escaping honest', () => {
  const xml = '<root a="1"><row r="1"><c r="A1" s="2"/></row><row r="2"/></root>';
  const row = findElement(xml, 'row');
  assert.equal(xml.slice(row.start, row.end), '<row r="1"><c r="A1" s="2"/></row>');
  assert.equal(findElements(xml, 'row').length, 2);
  assert.deepEqual(topLevelNames(xml), ['row', 'row']);

  assert.equal(getAttr(row.attrs, 'r'), '1');
  assert.equal(getAttr(row.attrs, 'missing'), null);
  assert.equal(setAttr(row.attrs, 'r', '7'), ' r="7"');
  assert.equal(setAttr(' r="1"', 'spans', '1:3'), ' r="1" spans="1:3"');

  const nasty = `<&"'>`;
  assert.equal(escapeXml(nasty), '&lt;&amp;&quot;&apos;&gt;');
  assert.equal(unescapeXml(escapeXml(nasty)), nasty);
  assert.equal(unescapeXml('&#1043;&#x41;'), 'ГA');
});

test('sheet relationship ids are resolved by namespace, not a conflicting r prefix', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  const book = DECODER.decode(source['xl/workbook.xml'])
    .replace('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
      'xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:r="urn:unrelated"')
    .replace(/r:id="([^"]+)"/g, 'rel:id="$1" r:id="wrong"');
  source['xl/workbook.xml'] = ENCODER.encode(book);
  const result = applyEdits(entriesFrom(zipSync(source)), [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 42 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.match(cellXml(sheetPart(result.entries, RICH_SUMMARY), 'B2'), /<v>42<\/v>/);
});

test('undeclared or ambiguous sheet relationship ids are refused without changes', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  const original = DECODER.decode(source['xl/workbook.xml']);
  for (const book of [
    original.replace(/r:id="([^"]+)"/, 'unknown:id="$1"'),
    original.replace(/r:id="([^"]+)"/, 'r:id="$1" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:id="$1"'),
  ]) {
    source['xl/workbook.xml'] = ENCODER.encode(book);
    const entries = entriesFrom(zipSync(source));
    const before = entries.map((entry) => entry.data.slice());
    const result = applyEdits(entries, [
      { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 42 } },
    ]);
    assert.equal(result.status, 'refused');
    assert.deepEqual(entries.map((entry) => entry.data), before);
  }
});

test('sheet relationship prefix overrides at sheet scope are honored', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/workbook.xml'] = ENCODER.encode(DECODER.decode(source['xl/workbook.xml'])
    .replace('name="Summary"', 'xmlns:r="urn:unrelated" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:id="rId1" name="Summary"')
    .replace(/(name="Summary"[^>]*?)r:id="rId1"/, '$1r:id="wrong"'));
  const result = applyEdits(entriesFrom(zipSync(source)), [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 42 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.match(cellXml(sheetPart(result.entries, RICH_SUMMARY), 'B2'), /<v>42<\/v>/);
});

test('conflicting workbook relationship IDs are refused before writing a cell', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  const path = 'xl/_rels/workbook.xml.rels';
  const rels = DECODER.decode(source[path]);
  const first = findElements(rels, 'Relationship').find((el) => getAttr(el.attrs, 'Id') === 'rId1');
  assert.ok(first);
  source[path] = ENCODER.encode(rels.replace('</Relationships>', `${rels.slice(first.start, first.end)}</Relationships>`));
  const entries = entriesFrom(zipSync(source));
  const before = entries.map((entry) => entry.data.slice());
  const result = applyEdits(entries, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 42 } },
  ]);
  assert.equal(result.status, 'refused');
  assert.deepEqual(entries.map((entry) => entry.data), before);
});

test('number edit touches only the cell and leaves every other entry byte-identical', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 424242 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const changed = changedNames(before, result.entries);
  // rich.xlsx already ships calcPr fullCalcOnLoad="1" and has no calcChain, so the only
  // part that has to change is the edited worksheet. Re-encoding the workbook just to
  // restamp an attribute it already has would be pointless churn.
  assert.deepEqual(changed, [RICH_SUMMARY]);

  const files = unpack(result.entries);
  assert.deepEqual(Object.keys(files).sort(), Object.keys(unzipSync(loadFixture('rich.xlsx'))).sort());

  const sheet = DECODER.decode(files[RICH_SUMMARY]);
  const b2 = cellXml(sheet, 'B2');
  assert.ok(b2, 'B2 still present');
  assert.match(b2, /<v>424242<\/v>/);
  assert.match(b2, / s="3"/, 'existing style kept');
  assert.doesNotMatch(b2, /t="inlineStr"/);

  // Every other entry, byte for byte, against the source archive.
  const original = unzipSync(loadFixture('rich.xlsx'));
  for (const [name, bytes] of Object.entries(files)) {
    if (name === RICH_SUMMARY || name === 'xl/workbook.xml') continue;
    assert.equal(Buffer.compare(Buffer.from(bytes), Buffer.from(original[name])), 0, `${name} changed`);
  }
  assert.equal(
    Buffer.compare(Buffer.from(files['xl/styles.xml']), Buffer.from(original['xl/styles.xml'])),
    0,
    'styles.xml untouched',
  );

  // Untouched cells inside the edited part keep their exact source text.
  const sourceSheet = DECODER.decode(original[RICH_SUMMARY]);
  for (const ref of ['A2', 'C2', 'D2', 'E2', 'A1', 'A9', 'B7', 'C7']) {
    assert.equal(cellXml(sheet, ref), cellXml(sourceSheet, ref), `${ref} must not move`);
  }
  assert.match(sheet, /<mergeCells count="1"><mergeCell ref="A9:E9" \/><\/mergeCells>/);
  assert.match(sheet, /<conditionalFormatting sqref="D2:D6">/);
  assert.match(sheet, /<dimension ref="A1:E9" \/>/, 'dimension untouched, original formatting kept');

  const book = DECODER.decode(files['xl/workbook.xml']);
  assert.match(book, /<calcPr[^>]*fullCalcOnLoad="1"/);

  const next = byName(result.entries);
  assert.equal(next.get(RICH_SUMMARY).changed, true);
  assert.equal(next.get(RICH_SUMMARY) === byName(before).get(RICH_SUMMARY), false, 'edited part is a new object');
  assert.equal(next.get('xl/workbook.xml').changed, false);
  assert.equal(next.get('xl/workbook.xml') === byName(before).get('xl/workbook.xml'), true, 'untouched parts are reused');
  assert.equal(next.get('xl/styles.xml').changed, false);
});

test('text values are inlined and survive a re-read', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const value = `<&"'> & амперсанд 100%`;
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'A2', value: { kind: 'text', value } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  const sheet = DECODER.decode(files[RICH_SUMMARY]);
  const a2 = cellXml(sheet, 'A2');
  assert.match(a2, /t="inlineStr"/);
  assert.match(a2, /<is><t xml:space="preserve">/);
  assert.match(a2, /&lt;&amp;&quot;&apos;&gt;/);
  assert.doesNotMatch(a2, /<&/, 'raw markup must never be emitted');

  const inner = /<is><t xml:space="preserve">([\s\S]*?)<\/t><\/is>/.exec(a2)[1];
  assert.equal(unescapeXml(inner), value);
  assert.equal(
    Buffer.compare(
      Buffer.from(files['xl/sharedStrings.xml'] ?? new Uint8Array()),
      Buffer.from(unzipSync(loadFixture('rich.xlsx'))['xl/sharedStrings.xml'] ?? new Uint8Array()),
    ),
    0,
    'sharedStrings untouched (and absent in this workbook)',
  );
});

test('formula writes f and cached v, drops calcChain, forces recalculation', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/calcChain.xml'] = ENCODER.encode(
    '<?xml version="1.0" encoding="UTF-8"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>',
  );
  source['[Content_Types].xml'] = ENCODER.encode(
    DECODER.decode(source['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    ),
  );
  source['xl/_rels/workbook.xml.rels'] = ENCODER.encode(
    DECODER.decode(source['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml" Id="rId6"/></Relationships>',
    ),
  );
  // Drop the pre-existing fullCalcOnLoad so the writer has to add it.
  source['xl/workbook.xml'] = ENCODER.encode(
    DECODER.decode(source['xl/workbook.xml']).replace(' fullCalcOnLoad="1"', ''),
  );

  const before = entriesFrom(zipSync(source, { level: 6 }));
  const result = applyEdits(before, [
    {
      sheet: 'Summary',
      ref: 'C2',
      value: { kind: 'formula', formula: 'SUM(B2:B6)*2', cached: 1234.5 },
    },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  assert.equal(files['xl/calcChain.xml'], undefined, 'calcChain removed');

  const contentTypes = DECODER.decode(files['[Content_Types].xml']);
  assert.doesNotMatch(contentTypes, /calcChain/);

  const rels = DECODER.decode(files['xl/_rels/workbook.xml.rels']);
  assert.doesNotMatch(rels, /calcChain/);

  const book = DECODER.decode(files['xl/workbook.xml']);
  assert.match(book, /<calcPr[^>]*fullCalcOnLoad="1"/);
  assert.match(book, /calcId="124519"/, 'other calcPr attributes survive');

  const c2 = cellXml(DECODER.decode(files[RICH_SUMMARY]), 'C2');
  assert.equal(c2, '<c r="C2" s="4"><f>SUM(B2:B6)*2</f><v>1234.5</v></c>');

  const changed = changedNames(before, result.entries);
  assert.deepEqual(changed, [
    '-xl/calcChain.xml',
    '[Content_Types].xml',
    'xl/_rels/workbook.xml.rels',
    'xl/workbook.xml',
    RICH_SUMMARY,
  ]);
});

test('a cell outside the old dimension widens it and gets no invented style', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'D5', value: { kind: 'number', value: 7 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  const sheet = DECODER.decode(files['xl/worksheets/sheet2.xml']);
  assert.match(sheet, /<dimension ref="A1:D5"\/>/);
  const d5 = cellXml(sheet, 'D5');
  assert.equal(d5, '<c r="D5"><v>7</v></c>');
  assert.doesNotMatch(d5, / s="/);

  const rows = findElements(sheet, 'row').map((el) => Number(getAttr(el.attrs, 'r')));
  assert.deepEqual(rows, [1, 3, 5], 'rows stay in ascending order');
  assert.match(sheet, /<row r="3">[\s\S]*?<\/row><row r="5">/, 'row 5 inserted after row 3');
});

test('a new cell can copy the style of the row it lands in when asked', () => {
  const before = entriesFrom(loadFixture('styled.xlsx'));
  const result = applyEdits(before, [
    {
      sheet: 'Ledger',
      ref: 'Q2',
      value: { kind: 'number', value: 1 },
      styleFrom: 'P2',
    },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet1.xml']);
  // P2 carries style 4 (banded fill + thin border + '#,##0'), so the new Q2 must ask for the same.
  assert.match(cellXml(sheet, 'Q2'), /^<c r="Q2" s="4"><v>1<\/v><\/c>$/);
  assert.match(sheet, /<dimension ref="A1:Q166"\/>|<dimension ref="A1:Q166"\/>/);
});

test('blank clears the value but keeps the style of an existing cell', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'blank' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  const b2 = cellXml(DECODER.decode(files[RICH_SUMMARY]), 'B2');
  assert.equal(b2, '<c r="B2" s="3"/>');
  assert.equal(Buffer.compare(Buffer.from(files['xl/styles.xml']), Buffer.from(unzipSync(loadFixture('rich.xlsx'))['xl/styles.xml'])), 0);

  // A blank over a styled formula cell drops t and keeps s too.
  const second = applyEdits(before, [{ sheet: 'Summary', ref: 'C2', value: { kind: 'blank' } }]);
  assert.equal(second.status, 'ok', second.reason);
  const c2 = cellXml(DECODER.decode(unpack(second.entries)[RICH_SUMMARY]), 'C2');
  assert.equal(c2, '<c r="C2" s="4"/>');
});

test('boolean writes t="b" with 1/0', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'B2', value: { kind: 'boolean', value: true } },
    { sheet: 'Notes', ref: 'B3', value: { kind: 'boolean', value: false } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']);
  assert.equal(cellXml(sheet, 'B2'), '<c r="B2" t="b"><v>1</v></c>');
  assert.equal(cellXml(sheet, 'B3'), '<c r="B3" t="b"><v>0</v></c>');
  const rows = findElements(sheet, 'row').map((el) => Number(getAttr(el.attrs, 'r')));
  assert.deepEqual(rows, [1, 2, 3], 'row 2 inserted before row 3');
});

test('date writes the bare serial, keeping the format and dropping the old t', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    // E2 already carries the yyyy-mm-dd format; 45306 is 2024-01-15 in the 1900 system.
    { sheet: 'Summary', ref: 'E2', value: { kind: 'date', value: 45306 } },
    // A fresh cell gets no format and no t: the caller chose the serial, the writer adds nothing.
    { sheet: 'Notes', ref: 'B5', value: { kind: 'date', value: 45306.5 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  const summary = DECODER.decode(files[RICH_SUMMARY]);
  const e2 = cellXml(summary, 'E2');
  assert.equal(e2, '<c r="E2" s="6"><v>45306</v></c>', 'serial only: no t, no text, style kept');
  assert.doesNotMatch(e2, /t="/);

  const notes = DECODER.decode(files['xl/worksheets/sheet2.xml']);
  assert.equal(cellXml(notes, 'B5'), '<c r="B5"><v>45306.5</v></c>', 'time of day stays fractional');
});

test('error writes t="e" and accepts only the codes Excel itself uses', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const codes = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];
  const result = applyEdits(
    before,
    codes.map((code, index) => ({
      sheet: 'Notes',
      ref: `B${index + 2}`,
      value: { kind: 'error', value: code },
    })),
  );
  assert.equal(result.status, 'ok', result.reason);

  const sheet = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']);
  for (const [index, code] of codes.entries()) {
    assert.equal(cellXml(sheet, `B${index + 2}`), `<c r="B${index + 2}" t="e"><v>${code}</v></c>`, code);
  }

  for (const bad of ['#NOPE!', 'DIV/0!', '#div/0!', '']) {
    const refused = applyEdits(before, [{ sheet: 'Summary', ref: 'B2', value: { kind: 'error', value: bad } }]);
    assert.equal(refused.status, 'refused', JSON.stringify(bad));
    assert.match(refused.reason, /not an Excel error code/);
  }
});

test('recalculation bookkeeping costs exactly the sheet, workbook and content types', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/workbook.xml'] = ENCODER.encode(
    DECODER.decode(source['xl/workbook.xml']).replace(' fullCalcOnLoad="1"', ''),
  );
  source['xl/calcChain.xml'] = ENCODER.encode(
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>',
  );
  source['[Content_Types].xml'] = ENCODER.encode(
    DECODER.decode(source['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    ),
  );

  const before = entriesFrom(zipSync(source, { level: 6 }));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 5 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  // Three parts change: the worksheet, workbook.xml (calc flag) and [Content_Types].xml
  // (calcChain override dropped). No calcChain relationship existed, so no rels rewrite.
  assert.deepEqual(changedNames(before, result.entries), [
    '-xl/calcChain.xml',
    '[Content_Types].xml',
    'xl/workbook.xml',
    RICH_SUMMARY,
  ]);
  assert.equal(result.notes.filter((note) => /calcChain/.test(note)).length, 2);
  assert.match(DECODER.decode(unpack(result.entries)['xl/workbook.xml']), /fullCalcOnLoad="1"/);
});

test('several cells in one row are all written in column order', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'D2', value: { kind: 'number', value: 0.5 } },
    { sheet: 'Summary', ref: 'A2', value: { kind: 'text', value: 'renamed' } },
    { sheet: 'Summary', ref: 'B5', value: { kind: 'number', value: 1 } },
    { sheet: 'Summary', ref: 'C2', value: { kind: 'blank' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = DECODER.decode(unpack(result.entries)[RICH_SUMMARY]);
  const row2 = findElements(sheet, 'row').find((el) => getAttr(el.attrs, 'r') === '2');
  const cols = findElements(sheet.slice(row2.start, row2.end), 'c').map((el) => getAttr(el.attrs, 'r'));
  assert.deepEqual(cols, ['A2', 'B2', 'C2', 'D2', 'E2'], 'row 2 keeps its cell order');
  assert.equal(cellXml(sheet, 'D2'), '<c r="D2" s="5"><v>0.5</v></c>');
  assert.equal(cellXml(sheet, 'C2'), '<c r="C2" s="4"/>');
  assert.match(cellXml(sheet, 'A2'), /renamed/);
  assert.equal(cellXml(sheet, 'B2'), cellXml(DECODER.decode(unzipSync(loadFixture('rich.xlsx'))[RICH_SUMMARY]), 'B2'));
  assert.equal(cellXml(sheet, 'B5'), '<c r="B5" s="8"><v>1</v></c>');
});

test('new cells in one empty row keep ascending column order', () => {
  const before = withSheet(sheetDocument('<row r="1"></row>', 'A1:E1'));
  const result = applyEdits(before, [
    // Reverse order on purpose: the row, not the edit list, decides the column order.
    { sheet: 'Notes', ref: 'E1', value: { kind: 'number', value: 5 } },
    { sheet: 'Notes', ref: 'C1', value: { kind: 'text', value: 'c' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = notesOf(result);
  assert.deepEqual(cellsInRow(sheet, 1), ['C1', 'E1'], 'two cells must not share one offset and reverse');
  assert.equal(cellXml(sheet, 'C1'), '<c r="C1" t="inlineStr"><is><t xml:space="preserve">c</t></is></c>');
  assert.equal(cellXml(sheet, 'E1'), '<c r="E1"><v>5</v></c>');
});

test('new cells between existing cells keep the row ascending', () => {
  const before = withSheet(
    sheetDocument('<row r="1"><c r="A1"><v>1</v></c><c r="D1"><v>4</v></c></row>', 'A1:D1'),
  );
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'C1', value: { kind: 'number', value: 3 } },
    { sheet: 'Notes', ref: 'B1', value: { kind: 'number', value: 2 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = notesOf(result);
  assert.deepEqual(cellsInRow(sheet, 1), ['A1', 'B1', 'C1', 'D1']);
  assert.equal(cellXml(sheet, 'A1'), '<c r="A1"><v>1</v></c>', 'existing cell stays as it was');
  assert.equal(cellXml(sheet, 'B1'), '<c r="B1"><v>2</v></c>');
  assert.equal(cellXml(sheet, 'C1'), '<c r="C1"><v>3</v></c>');
});

test('nine interleaved edits across three rows stay ordered', () => {
  const before = withSheet(
    sheetDocument(
      '<row r="1"><c r="B1"><v>0</v></c><c r="F1"><v>0</v></c></row>' +
        '<row r="2"></row>' +
        '<row r="3"><c r="A3"><v>0</v></c><c r="C3"><v>0</v></c></row>',
      'A1:F3',
    ),
  );
  const plan = [
    ['A1', 1],
    ['D1', 4],
    ['E1', 5],
    ['A2', 1],
    ['C2', 3],
    ['E2', 5],
    ['B3', 2],
    ['D3', 4],
    ['E3', 5],
  ];
  const result = applyEdits(
    before,
    plan.map(([ref, value]) => ({ sheet: 'Notes', ref, value: { kind: 'number', value } })),
  );
  assert.equal(result.status, 'ok', result.reason);

  const sheet = notesOf(result);
  assert.deepEqual(cellsInRow(sheet, 1), ['A1', 'B1', 'D1', 'E1', 'F1']);
  assert.deepEqual(cellsInRow(sheet, 2), ['A2', 'C2', 'E2']);
  assert.deepEqual(cellsInRow(sheet, 3), ['A3', 'B3', 'C3', 'D3', 'E3']);
  assert.equal(cellXml(sheet, 'D1'), '<c r="D1"><v>4</v></c>');
  assert.equal(cellXml(sheet, 'E2'), '<c r="E2"><v>5</v></c>');
});

test('refuses a shared formula master and lets a slave become independent', () => {
  const body =
    '<row r="2"><c r="B2"><f t="shared" ref="B2:B10" si="0">A2*2</f><v>4</v></c></row>' +
    '<row r="3"><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>';
  const before = withSheet(sheetDocument(body, 'A1:B10'));
  const master = '<c r="B2"><f t="shared" ref="B2:B10" si="0">A2*2</f><v>4</v></c>';

  const blocked = applyEdits(before, [
    { sheet: 'Notes', ref: 'B2', value: { kind: 'number', value: 1 } },
  ]);
  assert.equal(blocked.status, 'refused');
  assert.match(blocked.reason, /shared formula master at B2; edit a slave or the whole range/);

  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'B3', value: { kind: 'number', value: 5 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = notesOf(result);
  assert.equal(cellXml(sheet, 'B2'), master, 'the master every slave pointed at is byte-identical');
  const b3 = cellXml(sheet, 'B3');
  assert.equal(b3, '<c r="B3"><v>5</v></c>');
  assert.doesNotMatch(b3, /si=|t="shared"/, 'the slave is no longer part of the shared formula');
});

test('cell text keeps literal _xHHHH_, controls and CR through Excel escapes', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'B2', value: { kind: 'text', value: '_x0041_ lit' } },
    { sheet: 'Notes', ref: 'B3', value: { kind: 'text', value: 'a\u0001b' } },
    { sheet: 'Notes', ref: 'B4', value: { kind: 'text', value: 'l1\r\nl2' } },
    { sheet: 'Notes', ref: 'B5', value: { kind: 'formula', formula: 'A1+_x0041_', cached: '_x0041_' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']);
  const written = (ref) => /<t xml:space="preserve">([\s\S]*?)<\/t>/.exec(cellXml(sheet, ref))[1];
  assert.equal(written('B2'), '_x005F_x0041_ lit', 'Excel would otherwise read the literal as "A"');
  assert.equal(written('B3'), 'a_x0001_b', 'XML cannot carry a raw C0 control');
  assert.equal(written('B4'), 'l1_x000D_\nl2', 'a raw CR would be normalised away by the parser');
  assert.equal(
    cellXml(sheet, 'B5'),
    '<c r="B5" t="str"><f>A1+_x0041_</f><v>_x005F_x0041_</v></c>',
    'a formula is not cell text; only its string result is escaped',
  );
});

test('text over 32767 characters is refused with the exact length', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const refused = applyEdits(before, [
    { sheet: 'Notes', ref: 'B2', value: { kind: 'text', value: 'x'.repeat(32768) } },
  ]);
  assert.equal(refused.status, 'refused');
  assert.match(refused.reason, /32768 characters/);
  assert.match(refused.reason, /32767-character cell limit/);

  const atLimit = applyEdits(before, [
    { sheet: 'Notes', ref: 'B2', value: { kind: 'text', value: 'y'.repeat(32767) } },
  ]);
  assert.equal(atLimit.status, 'ok', atLimit.reason);

  const cached = applyEdits(before, [
    { sheet: 'Notes', ref: 'B3', value: { kind: 'formula', formula: 'A1', cached: 'z'.repeat(40000) } },
  ]);
  assert.equal(cached.status, 'refused');
  assert.match(cached.reason, /40000 characters/);
});

test('refuses a non-anchor cell inside an array formula range', () => {
  const before = withSheet(
    sheetDocument(
      '<row r="5"><c r="C5"><f t="array" ref="C5:D6">B5:B6*2</f><v>1</v></c><c r="D5"><v>2</v></c></row>',
      'A1:D6',
    ),
  );

  const blocked = applyEdits(before, [
    { sheet: 'Notes', ref: 'D5', value: { kind: 'number', value: 9 } },
  ]);
  assert.equal(blocked.status, 'refused');
  assert.match(blocked.reason, /inside array formula range C5:D6/);
  assert.match(blocked.reason, /only C5 is writable/);

  const anchor = applyEdits(before, [
    { sheet: 'Notes', ref: 'C5', value: { kind: 'number', value: 9 } },
  ]);
  assert.equal(anchor.status, 'ok', anchor.reason);
  assert.equal(cellXml(notesOf(anchor), 'C5'), '<c r="C5"><v>9</v></c>');
});

test('reports untracked cm="1" spill markers instead of guessing ranges', () => {
  const before = withSheet(
    sheetDocument(
      '<row r="1"><c r="A1" cm="1"><f>SEQUENCE(3)</f><v>1</v></c></row>' +
        '<row r="2"><c r="A2"><v>2</v></c></row>',
      'A1:A3',
    ),
  );
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'A2', value: { kind: 'number', value: 5 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.ok(
    result.notes.some((note) => /cm="1" spill markers are present/.test(note)),
    JSON.stringify(result.notes),
  );
  assert.match(result.notes.join('\n'), /only t="array" ranges are protected/);
  assert.equal(cellXml(notesOf(result), 'A2'), '<c r="A2"><v>5</v></c>');
});

test('new cells land before a row-level extLst', () => {
  const before = withSheet(
    sheetDocument(
      '<row r="1"><c r="A1"><v>1</v></c>' +
        '<extLst><ext uri="{TEST}"><meta xmlns="urn:test" v="1"/></ext></extLst></row>',
      'A1:A1',
    ),
  );
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'A1', value: { kind: 'number', value: 11 } },
    { sheet: 'Notes', ref: 'Z1', value: { kind: 'number', value: 26 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = notesOf(result);
  const row = findElements(sheet, 'row').find((el) => getAttr(el.attrs, 'r') === '1');
  const rowText = sheet.slice(row.start, row.end);
  assert.match(rowText, /<c r="A1"><v>11<\/v><\/c><c r="Z1"><v>26<\/v><\/c><extLst>/);
  assert.match(rowText, /<meta xmlns="urn:test" v="1"\/>/, 'row metadata survives intact');
});

test('a prefixed worksheet root gets prefixed cells, rows and dimension', () => {
  const before = withSheet(
    '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<x:dimension ref="A1:A1"/><x:sheetData>' +
      '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>' +
      '</x:sheetData></x:worksheet>',
  );
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'A1', value: { kind: 'text', value: 'hello' } },
    { sheet: 'Notes', ref: 'C2', value: { kind: 'number', value: 2 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = notesOf(result);
  assert.equal(
    cellXml(sheet, 'A1'),
    '<x:c r="A1" t="inlineStr"><x:is><x:t xml:space="preserve">hello</x:t></x:is></x:c>',
  );
  assert.match(sheet, /<x:row r="2"><x:c r="C2"><x:v>2<\/x:v><\/x:c><\/x:row>/);
  assert.match(sheet, /<x:dimension ref="A1:C2"\/>/);
  assert.doesNotMatch(
    sheet,
    /<(c|row|dimension|is|t|v|f)[\s/>]/,
    'an unprefixed element would fall out of the spreadsheetml namespace',
  );
});

test('a cell edited twice resolves to the last value', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 1 } },
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 2 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.equal(cellXml(DECODER.decode(unpack(result.entries)[RICH_SUMMARY]), 'B2'), '<c r="B2" s="3"><v>2</v></c>');
});

test('edits in later rows leave earlier rows byte-identical', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const original = DECODER.decode(unzipSync(loadFixture('rich.xlsx'))[RICH_SUMMARY]);
  const rowsBefore = (xml) => {
    const sd = findElement(xml, 'sheetData');
    const inner = xml.slice(sd.innerStart, sd.innerEnd);
    const map = new Map();
    for (const row of findElements(inner, 'row')) map.set(getAttr(row.attrs, 'r'), inner.slice(row.start, row.end));
    return map;
  };

  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B4', value: { kind: 'number', value: 1 } },
    { sheet: 'Summary', ref: 'E6', value: { kind: 'text', value: 'x' } },
    { sheet: 'Summary', ref: 'A10', value: { kind: 'number', value: 1 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const after = rowsBefore(DECODER.decode(unpack(result.entries)[RICH_SUMMARY]));
  const source = rowsBefore(original);
  for (const ref of ['1', '2', '3', '5', '7', '9']) {
    assert.equal(after.get(ref), source.get(ref), `row ${ref} untouched`);
  }
  assert.deepEqual([...after.keys()], ['1', '2', '3', '4', '5', '6', '7', '9', '10']);
  assert.match(cellXml(DECODER.decode(unpack(result.entries)[RICH_SUMMARY]), 'A10'), /<v>1<\/v>/);
});

test('blanking a cell that never existed changes nothing at all', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'Z99', value: { kind: 'blank' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.deepEqual(changedNames(before, result.entries), [], 'no part rewritten');
  assert.equal(result.entries.length, before.length);
  assert.match(DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']), /<dimension ref="A1:A3" \/>/);
});

test('an empty edit list is a no-op that still returns a fresh array', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, []);
  assert.equal(result.status, 'ok');
  assert.notEqual(result.entries, before, 'new array');
  assert.deepEqual(result.entries, before, 'same entries');
});

test('xml.js finds a top-level element even when a nested one shares its name', () => {
  const xml =
    '<worksheet><sheetPr><pageSetUpPr/></sheetPr><dimension ref="A1:B2"/><sheetData><row r="1"/></sheetData></worksheet>';
  assert.deepEqual(topLevelNames(xml), ['sheetPr', 'dimension', 'sheetData']);
  const nested = '<root><extLst><extLst/></extLst><extLst a="1"/></root>';
  assert.deepEqual(topLevelNames(nested), ['extLst', 'extLst']);
  const first = findElement(nested, 'extLst');
  assert.equal(nested.slice(first.start, first.end), '<extLst><extLst/></extLst>');
});

test('refuses malformed edits before touching any part', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const cases = [
    [{ sheet: 'Summary', ref: 'not-a-ref', value: { kind: 'number', value: 1 } }, /invalid cell reference/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'nope' } }, /unsupported value kind/],
    [{ sheet: '', ref: 'A1', value: { kind: 'blank' } }, /no sheet name/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'number', value: Number.NaN } }, /not a finite number/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'date', value: Number.NaN } }, /finite Excel serial/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'date', value: new Date() } }, /finite Excel serial/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'date', value: -1 } }, /outside the representable range/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'date', value: 2958466 } }, /outside the representable range/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'error', value: '#NOPE!' } }, /not an Excel error code/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'error', value: 7 } }, /not an Excel error code/],
    [{ sheet: 'Summary', ref: 'A1', value: { kind: 'formula', formula: '  ' } }, /formula is empty/],
    [{ sheet: 'Summary', ref: 'A1' }, /no value/],
  ];
  for (const [edit, reason] of cases) {
    const result = applyEdits(before, [edit]);
    assert.equal(result.status, 'refused', JSON.stringify(edit));
    assert.match(result.reason, reason);
  }
});

test('xlsxPackage resolves sheets through the workbook relationships', () => {
  const entries = entriesFrom(loadFixture('rich.xlsx'));
  const pkg = xlsxPackage(entries);
  assert.ok(pkg);
  assert.equal(pkg.book.path, 'xl/workbook.xml');
  assert.deepEqual(
    pkg.book.sheets.map((sheet) => [sheet.name, sheet.part]),
    [
      ['Summary', 'xl/worksheets/sheet1.xml'],
      ['Notes', 'xl/worksheets/sheet2.xml'],
      ['Plain', 'xl/worksheets/sheet3.xml'],
    ],
  );
  assert.ok(pkg.files instanceof Map);
  assert.equal(pkg.files.get('xl/workbook.xml'), entries.find((entry) => entry.name === 'xl/workbook.xml'));
  assert.equal(xlsxPackage([]), null, 'an empty package has no workbook');
});

test('a package with relative rel targets and a foreign r: prefix still resolves', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/_rels/workbook.xml.rels'] = ENCODER.encode(
    DECODER.decode(source['xl/_rels/workbook.xml.rels']).replace(/Target="\/xl\/worksheets\//g, 'Target="worksheets/'),
  );
  source['xl/workbook.xml'] = ENCODER.encode(
    DECODER.decode(source['xl/workbook.xml'])
      .replace('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"', 'xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
      .replace(/r:id=/g, 'rel:id='),
  );
  const before = entriesFrom(zipSync(source, { level: 6 }));

  const pkg = xlsxPackage(before);
  assert.equal(pkg.book.sheets[0].part, 'xl/worksheets/sheet1.xml', 'relative target resolved');

  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 8 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.equal(cellXml(DECODER.decode(unpack(result.entries)[RICH_SUMMARY]), 'B2'), '<c r="B2" s="3"><v>8</v></c>');
});

test('a calc chain declared only through a relationship is removed too', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/metrics/recalc.xml'] = ENCODER.encode(
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>',
  );
  source['xl/_rels/workbook.xml.rels'] = ENCODER.encode(
    DECODER.decode(source['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="metrics/recalc.xml" Id="rId6"/></Relationships>',
    ),
  );
  const before = entriesFrom(zipSync(source, { level: 6 }));

  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 3 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  assert.equal(files['xl/metrics/recalc.xml'], undefined, 'chain found through its relationship');
  assert.doesNotMatch(DECODER.decode(files['xl/_rels/workbook.xml.rels']), /calcChain/);
});

test('an empty sheetData gains a row, and a missing dimension is created', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  const sheet = DECODER.decode(source['xl/worksheets/sheet3.xml']);
  const stripped = sheet
    .replace(/<dimension[^>]*\/>/, '')
    .replace(/<sheetData>[\s\S]*<\/sheetData>/, '<sheetData></sheetData>');
  source['xl/worksheets/sheet3.xml'] = ENCODER.encode(stripped);
  const before = entriesFrom(zipSync(source, { level: 6 }));

  const result = applyEdits(before, [
    { sheet: 'Plain', ref: 'B3', value: { kind: 'text', value: 'first' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const out = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet3.xml']);
  assert.match(out, /<dimension ref="B3"\/>/, 'dimension added, single cell written without a range');
  assert.match(
    out,
    /<sheetData><row r="3"><c r="B3" t="inlineStr"><is><t xml:space="preserve">first<\/t><\/is><\/c><\/row><\/sheetData>/,
  );
  assert.ok(out.indexOf('<dimension') < out.indexOf('<sheetData'), 'dimension precedes sheetData');
});

test('a self-closing row receives its first cell', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/worksheets/sheet2.xml'] = ENCODER.encode(
    DECODER.decode(source['xl/worksheets/sheet2.xml']).replace(/<row r="1">[\s\S]*?<\/row>/, '<row r="1" ht="12" />'),
  );
  const before = entriesFrom(zipSync(source, { level: 6 }));

  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'A1', value: { kind: 'number', value: 11 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const out = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']);
  assert.match(out, /<row r="1" ht="12"><c r="A1"><v>11<\/v><\/c><\/row>/, 'row attributes survive reopening');
});

test('a shared-string cell becomes a plain value without leaving t="s"', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/worksheets/sheet2.xml'] = ENCODER.encode(
    DECODER.decode(source['xl/worksheets/sheet2.xml']).replace(
      '<c r="A1" t="inlineStr"><is><t>Sheet with no styling</t></is></c>',
      '<c r="A1" s="3" t="s"><v>0</v></c>',
    ),
  );
  const before = entriesFrom(zipSync(source, { level: 6 }));

  const numeric = applyEdits(before, [
    { sheet: 'Notes', ref: 'A1', value: { kind: 'number', value: 42 } },
  ]);
  assert.equal(numeric.status, 'ok', numeric.reason);
  assert.equal(
    cellXml(DECODER.decode(unpack(numeric.entries)['xl/worksheets/sheet2.xml']), 'A1'),
    '<c r="A1" s="3"><v>42</v></c>',
    'shared-string marker dropped so the index is no longer read as a table lookup',
  );

  const textual = applyEdits(before, [
    { sheet: 'Notes', ref: 'A1', value: { kind: 'text', value: 'inline now' } },
  ]);
  assert.equal(textual.status, 'ok', textual.reason);
  assert.equal(
    cellXml(DECODER.decode(unpack(textual.entries)['xl/worksheets/sheet2.xml']), 'A1'),
    '<c r="A1" s="3" t="inlineStr"><is><t xml:space="preserve">inline now</t></is></c>',
  );
});

test('a formula without a cached value drops the stale <v>', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B7', value: { kind: 'formula', formula: 'SUM(B2:B6)' } },
    { sheet: 'Summary', ref: 'C7', value: { kind: 'formula', formula: 'SUM(C2:C6)', cached: null } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  const sheet = DECODER.decode(unpack(result.entries)[RICH_SUMMARY]);
  assert.equal(cellXml(sheet, 'B7'), '<c r="B7" s="13"><f>SUM(B2:B6)</f></c>');
  // The model reports a formula without a cached result as `cached: null`; that must never be
  // written as a <v>, or the cell would claim an empty string as its result.
  assert.equal(cellXml(sheet, 'C7'), '<c r="C7" s="14"><f>SUM(C2:C6)</f></c>', 'explicit null is not a value');
});

test('a formula cached as a boolean or a string carries the right t', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'C1', value: { kind: 'formula', formula: 'ISBLANK(A1)', cached: false } },
    { sheet: 'Notes', ref: 'C2', value: { kind: 'formula', formula: 'A1&""', cached: 'text <result>' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const sheet = DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']);
  assert.equal(cellXml(sheet, 'C1'), '<c r="C1" t="b"><f>ISBLANK(A1)</f><v>0</v></c>');
  assert.equal(
    cellXml(sheet, 'C2'),
    '<c r="C2" t="str"><f>A1&amp;&quot;&quot;</f><v>text &lt;result&gt;</v></c>',
    'formula text is escaped, string result keeps its own t',
  );
  // Quotes may be escaped or literal in element text; the reader must see the source back.
  const inner = /<f>([\s\S]*?)<\/f>/.exec(cellXml(sheet, 'C2'))[1];
  assert.equal(unescapeXml(inner), 'A1&""');
});

test('sheet names are matched case-insensitively', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'summary', ref: 'b2', value: { kind: 'number', value: 99 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.equal(cellXml(DECODER.decode(unpack(result.entries)[RICH_SUMMARY]), 'B2'), '<c r="B2" s="3"><v>99</v></c>');
});

test('macro.xlsm keeps vbaProject.bin and styles.xml byte-identical', () => {
  const original = loadFixture('macro.xlsm');
  const before = entriesFrom(original);
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 31337 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  const source = unzipSync(original);
  for (const name of ['xl/vbaProject.bin', 'xl/styles.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']) {
    assert.equal(Buffer.compare(Buffer.from(files[name]), Buffer.from(source[name])), 0, `${name} changed`);
  }
  assert.match(DECODER.decode(files[RICH_SUMMARY]), /<v>31337<\/v>/);
  assert.deepEqual(changedNames(before, result.entries), [RICH_SUMMARY]);
});

test('refuses a missing sheet', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Nope', ref: 'A1', value: { kind: 'number', value: 1 } },
  ]);
  assert.equal(result.status, 'refused');
  assert.match(result.reason, /not found/i);
  assert.equal(result.entries, undefined);
});

test('refuses a cell inside a merged range but allows its anchor', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const blocked = applyEdits(before, [
    { sheet: 'Summary', ref: 'C9', value: { kind: 'text', value: 'nope' } },
  ]);
  assert.equal(blocked.status, 'refused');
  assert.match(blocked.reason, /merged range A9:E9/);

  const allowed = applyEdits(before, [
    { sheet: 'Summary', ref: 'A9', value: { kind: 'text', value: 'anchor' } },
  ]);
  assert.equal(allowed.status, 'ok', allowed.reason);
  const a9 = cellXml(DECODER.decode(unpack(allowed.entries)[RICH_SUMMARY]), 'A9');
  assert.match(a9, /<t xml:space="preserve">anchor<\/t>/);
  assert.match(a9, / s="15"/);
});

test('refuses a password-protected sheet', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/worksheets/sheet2.xml'] = ENCODER.encode(
    DECODER.decode(source['xl/worksheets/sheet2.xml']).replace(
      '<sheetFormatPr',
      '<sheetProtection sheet="1" objects="1" scenarios="1" algorithmName="SHA-512" hashValue="x"/><sheetFormatPr',
    ),
  );
  const before = entriesFrom(zipSync(source, { level: 6 }));
  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'A1', value: { kind: 'text', value: 'x' } },
  ]);
  assert.equal(result.status, 'refused');
  assert.match(result.reason, /protected/i);
});

test('refuses a part that is not a worksheet', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const broken = before.map((entry) =>
    entry.name === RICH_SUMMARY
      ? { ...entry, data: ENCODER.encode('<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>') }
      : entry,
  );
  const result = applyEdits(broken, [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 1 } },
  ]);
  assert.equal(result.status, 'refused');
  assert.match(result.reason, /not a worksheet/);
});

test('refuses a broken package instead of guessing', () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const noWorkbook = before.filter((entry) => entry.name !== 'xl/workbook.xml');
  assert.equal(applyEdits(noWorkbook, [{ sheet: 'Summary', ref: 'A1', value: { kind: 'blank' } }]).status, 'refused');

  const noSheet = before.filter((entry) => entry.name !== RICH_SUMMARY);
  const result = applyEdits(noSheet, [{ sheet: 'Summary', ref: 'A1', value: { kind: 'blank' } }]);
  assert.equal(result.status, 'refused');
  assert.match(result.reason, /missing/i);
});

test('styled.xlsx keeps its entry count, valid XML and shared sharedStrings', () => {
  const original = loadFixture('styled.xlsx');
  const source = unzipSync(original);
  const sourceSheet = DECODER.decode(source['xl/worksheets/sheet1.xml']);

  // The fixture ships no shared strings; add one so "untouched" is a real assertion
  // rather than a check that two absentities match.
  const seeded = { ...source };
  seeded['xl/sharedStrings.xml'] = ENCODER.encode(
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>kept</t></si></sst>',
  );
  const before = entriesFrom(zipSync(seeded, { level: 6 }));
  const seededBytes = seeded['xl/sharedStrings.xml'];

  const result = applyEdits(before, [
    { sheet: 'Ledger', ref: 'A2', value: { kind: 'number', value: 4242 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);
  assert.equal(result.entries.length, before.length, 'entry count unchanged');

  const files = unpack(result.entries);
  assert.equal(Object.keys(files).length, Object.keys(before).length);
  assert.equal(
    Buffer.compare(Buffer.from(files['xl/styles.xml']), Buffer.from(source['xl/styles.xml'])),
    0,
  );
  assert.equal(
    Buffer.compare(Buffer.from(files['xl/sharedStrings.xml']), Buffer.from(seededBytes)),
    0,
    'sharedStrings.xml byte-identical',
  );

  const sheet = DECODER.decode(files['xl/worksheets/sheet1.xml']);
  assert.equal(cellXml(sheet, 'A2'), '<c r="A2" s="2"><v>4242</v></c>');
  assert.equal(countDepth(sheet), 0, 'every tag closes');
  assert.equal(countOpen(sheet), countClose(sheet), 'open and close tags balanced');
  assert.equal(sheet.length - sourceSheet.length, '<c r="A2" s="2"><v>4242</v></c>'.length - cellLength(sourceSheet, 'A2'));
  assert.equal(cellLength(sheet, 'A2'), '<c r="A2" s="2"><v>4242</v></c>'.length);
});

test('a shared-strings workbook keeps sharedStrings.xml untouched by a text edit', () => {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/sharedStrings.xml'] = ENCODER.encode(
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Shared</t></si></sst>',
  );
  const before = entriesFrom(zipSync(source, { level: 6 }));

  const result = applyEdits(before, [
    { sheet: 'Notes', ref: 'A2', value: { kind: 'text', value: 'Shared' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const files = unpack(result.entries);
  assert.equal(
    Buffer.compare(Buffer.from(files['xl/sharedStrings.xml']), Buffer.from(source['xl/sharedStrings.xml'])),
    0,
    'shared string table is never rewritten',
  );
  // The text still lands inline, so nothing points into the shared table.
  assert.match(cellXml(DECODER.decode(files['xl/worksheets/sheet2.xml']), 'A2'), /t="inlineStr"/);
});

test('a written date and error come back through the model with their own kinds', async () => {
  const before = entriesFrom(loadFixture('rich.xlsx'));
  const result = applyEdits(before, [
    { sheet: 'Summary', ref: 'E2', value: { kind: 'date', value: 45306 } },
    { sheet: 'Summary', ref: 'D2', value: { kind: 'error', value: '#DIV/0!' } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const model = await loadWorkbook(new Uint8Array(await zipToBuffer(result.entries)));
  assert.equal(model.date1904, false, 'rich.xlsx is a 1900-system book');

  const date = model.sheets[0].cells.get('1,4');
  assert.equal(date.value.kind, 'date');
  assert.equal(date.value.value.toISOString(), '2024-01-15T00:00:00.000Z');
  assert.equal(date.value.display, '2024-01-15');

  const error = model.sheets[0].cells.get('1,3');
  assert.equal(error.value.kind, 'error');
  assert.equal(error.value.value, '#DIV/0!');
  assert.equal(error.value.display, '#DIV/0!');
});

test('the edited package loads back through a real workbook reader', async () => {
  const result = applyEdits(entriesFrom(loadFixture('rich.xlsx')), [
    { sheet: 'Summary', ref: 'B2', value: { kind: 'number', value: 424242 } },
    { sheet: 'Summary', ref: 'D2', value: { kind: 'blank' } },
    { sheet: 'Notes', ref: 'B2', value: { kind: 'text', value: 'typed <&>' } },
    { sheet: 'Plain', ref: 'A2', value: { kind: 'formula', formula: '1+2', cached: 3 } },
  ]);
  assert.equal(result.status, 'ok', result.reason);

  const book = new ExcelJS.Workbook();
  await book.xlsx.load(Buffer.from(await zipToBuffer(result.entries)));

  assert.deepEqual(book.worksheets.map((sheet) => sheet.name), ['Summary', 'Notes', 'Plain']);
  assert.equal(book.getWorksheet('Summary').getCell('B2').value, 424242);
  assert.equal(book.getWorksheet('Summary').getCell('D2').value, null);
  assert.equal(book.getWorksheet('Summary').getCell('C2').formula, 'B2*48.7');
  assert.equal(book.getWorksheet('Notes').getCell('B2').value, 'typed <&>');
  assert.equal(book.getWorksheet('Plain').getCell('A2').formula, '1+2');
  // Styles that were already there still resolve to a real style, not the default.
  assert.equal(book.getWorksheet('Summary').getCell('B2').numFmt, '#,##0');
});

/** Minimal Notes worksheet around a `sheetData` body; the rest of rich.xlsx stays intact. */
function sheetDocument(body, dimension = 'A1:A3') {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="${dimension}"/><sheetData>${body}</sheetData></worksheet>`
  );
}

/** rich.xlsx whose Notes part is replaced by `document`, so a test can shape sheetData. */
function withSheet(document) {
  const source = unpack(entriesFrom(loadFixture('rich.xlsx')));
  source['xl/worksheets/sheet2.xml'] = ENCODER.encode(document);
  return entriesFrom(zipSync(source, { level: 6 }));
}

function notesOf(result) {
  return DECODER.decode(unpack(result.entries)['xl/worksheets/sheet2.xml']);
}

/** `r` attributes of one row's cells, in document order. */
function cellsInRow(xml, rowNum) {
  const row = findElements(xml, 'row').find((el) => getAttr(el.attrs, 'r') === String(rowNum));
  assert.ok(row, `row ${rowNum} is present`);
  return findElements(xml.slice(row.start, row.end), 'c').map((el) => getAttr(el.attrs, 'r'));
}

async function zipToBuffer(entries) {
  const files = {};
  for (const entry of entries) files[entry.name] = entry.data;
  return zipSync(files, { level: 6 });
}


function cellLength(xml, ref) {
  const el = findElements(xml, 'c').find((candidate) => getAttr(candidate.attrs, 'r') === ref);
  return el ? el.end - el.start : -1;
}

function countOpen(xml) {
  return (xml.match(/<[A-Za-z][^>]*?(?<!\/)>/g) ?? []).length;
}

function countClose(xml) {
  return (xml.match(/<\/[A-Za-z][^>]*>/g) ?? []).length;
}

function countDepth(xml) {
  const tagRe = /<(\/?)([A-Za-z][\w.:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
  let depth = 0;
  let min = 0;
  let m;
  while ((m = tagRe.exec(xml)) !== null) {
    if (m[4] === '/') continue;
    if (m[1] === '/') depth--;
    else depth++;
    min = Math.min(min, depth);
  }
  return min;
}
