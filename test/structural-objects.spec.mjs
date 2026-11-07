import assert from 'node:assert/strict';
import { test } from 'node:test';
import { rewriteStructuralObjects } from '../src/write/structural-objects.js';

const enc = new TextEncoder();
const dec = new TextDecoder();
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pkgRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const draw = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const sheetPath = 'xl/worksheets/custom.xml';
const tablePath = 'xl/tables/t.xml';
const drawingPath = 'xl/drawings/d.xml';
const table = `<table xmlns="${main}" id="1" name="T" displayName="T" ref="B2:C6"><autoFilter ref="B2:C6"/><tableColumns count="2"><tableColumn id="1" name="A"><calculatedColumnFormula>SUM(A3:A6)</calculatedColumnFormula></tableColumn><tableColumn id="2" name="B"/></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>`;
const drawing = `<xdr:wsDr xmlns:xdr="${draw}"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp/><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
function fixture(options = {}) {
  const parts = {
    [sheetPath]: `<worksheet xmlns="${main}" xmlns:q="${rel}"><tableParts count="1"><tablePart q:id="t"/></tableParts><drawing q:id="d"/></worksheet>`,
    'xl/worksheets/_rels/custom.xml.rels': `<Relationships xmlns="${pkgRel}"><Relationship Id="t" Type="${rel}/table" Target="../tables/t.xml"/><Relationship Id="d" Type="${rel}/drawing" Target="../drawings/d.xml"/></Relationships>`,
    [tablePath]: table, [drawingPath]: drawing, 'xl/media/image.png': 'untouched', ...options,
  };
  const entries = Object.entries(parts).map(([name, text]) => ({ name, data: enc.encode(text), changed: false }));
  return { entries, sheet: { name: 'Data', part: sheetPath, entry: entries.find(e => e.name === sheetPath) } };
}
function run(f, options = {}) {
  return rewriteStructuralObjects({ entries: f.entries }, f.sheet, { sheet: 'Data', axis: 'row', action: 'insert', index: 4, count: 2, ...options });
}
function text(result, path) { return result.replacements.get(path)?.text; }

test('inserting rows expands table and filter, translates formulas and moves drawing endpoints without touching the package', () => {
  const f = fixture();
  const before = f.entries.map(e => dec.decode(e.data));
  const result = run(f);
  assert.equal(result.status, 'ok');
  assert.match(text(result, tablePath), /ref="B2:C8"/);
  assert.match(text(result, tablePath), /<autoFilter ref="B2:C8"/);
  assert.match(text(result, tablePath), /SUM\(A3:A8\)/);
  assert.match(text(result, drawingPath), /<xdr:row>7<\/xdr:row>/);
  assert.match(text(result, tablePath), /<tableStyleInfo name="TableStyleMedium2"\/>/);
  assert.equal(result.replacements.has('xl/media/image.png'), false);
  assert.deepEqual(f.entries.map(e => dec.decode(e.data)), before);
});

test('deleting interior rows shrinks ranges and anchor endpoints', () => {
  const result = run(fixture(), { action: 'delete', index: 3, count: 2 });
  assert.equal(result.status, 'ok');
  assert.match(text(result, tablePath), /ref="B2:C4"/);
  assert.match(text(result, drawingPath), /<xdr:row>3<\/xdr:row>/);
});

test('insertions at the start move tables and insertions after the end leave their extents intact', () => {
  const moved = run(fixture(), { index: 2, count: 1 });
  assert.equal(moved.status, 'ok');
  assert.match(text(moved, tablePath), /ref="B3:C7"/);
  const after = run(fixture(), { index: 7, count: 1 });
  assert.equal(after.status, 'ok');
  assert.equal(after.replacements.has(tablePath), false);
});

test('refuses deleting the table header or changing its column definitions', () => {
  for (const op of [{ action: 'delete', index: 2, count: 1 }, { axis: 'column', index: 3, count: 1 }, { axis: 'column', action: 'delete', index: 2, count: 1 }]) {
    const f = fixture();
    const original = f.entries.map(e => dec.decode(e.data));
    assert.equal(run(f, op).status, 'refused');
    assert.deepEqual(f.entries.map(e => dec.decode(e.data)), original);
  }
});

test('refuses unresolved, external and unknown relationships before producing changes', () => {
  for (const relationships of [
    `<Relationship Id="t" Type="${rel}/table" Target="../tables/missing.xml"/>`,
    `<Relationship Id="t" Type="${rel}/table" Target="https://example.test/t.xml" TargetMode="External"/>`,
    `<Relationship Id="t" Type="${rel}/unknownObject" Target="../tables/t.xml"/>`,
  ]) {
    const f = fixture({ 'xl/worksheets/_rels/custom.xml.rels': `<Relationships xmlns="${pkgRel}">${relationships}</Relationships>` });
    assert.equal(run(f).status, 'refused');
  }
});

test('rejects charts rather than leaving their data references stale', () => {
  const f = fixture({ [drawingPath]: drawing.replace('<xdr:sp/>', '<xdr:graphicFrame/>') });
  assert.equal(run(f).status, 'refused');
});

test('cross-sheet table formulas update without moving that sheet’s object geometry', () => {
  const f = fixture({ [tablePath]: table.replace('SUM(A3:A6)', "SUM(Data!A3:A6)") });
  f.sheet.name = 'Other';
  const result = run(f);
  assert.equal(result.status, 'ok');
  assert.match(text(result, tablePath), /ref="B2:C6"/);
  assert.match(text(result, tablePath), /SUM\(Data!A3:A8\)/);
  assert.equal(result.replacements.has(drawingPath), false);
});

test('relationship IDs use their declared namespace, not a misleading prefix', () => {
  const f = fixture({ [sheetPath]: `<worksheet xmlns="${main}" xmlns:r="urn:wrong" xmlns:q="${rel}"><tableParts count="1"><tablePart r:id="wrong" q:id="t"/></tableParts><drawing q:id="d"/></worksheet>` });
  assert.equal(run(f).status, 'ok');
});

test('preserves BOM metadata and records visited image parts without rewriting them', () => {
  const picture = drawing.replace('<xdr:sp/>', `<xdr:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"><a:blip r:embed="img"/></xdr:pic>`);
  const f = fixture({ [drawingPath]: picture, 'xl/drawings/_rels/d.xml.rels': `<Relationships xmlns="${pkgRel}"><Relationship Id="img" Type="${rel}/image" Target="../media/image.png"/></Relationships>` });
  const entry = f.entries.find(e => e.name === tablePath);
  entry.data = Uint8Array.from([0xef, 0xbb, 0xbf, ...entry.data]);
  const result = run(f);
  assert.equal(result.status, 'ok');
  assert.equal(result.replacements.get(tablePath).bom, true);
  assert.ok(result.visited.has('xl/media/image.png'));
  assert.ok(result.visited.has('xl/drawings/_rels/d.xml.rels'));
  assert.equal(result.replacements.has('xl/media/image.png'), false);
});
