import assert from 'node:assert/strict';
import { test } from 'node:test';
import { rewriteStructuralFormula } from '../src/write/structural-formula.js';

const base = { sheet: 'Data', formulaSheet: 'Data', axis: 'row', action: 'insert', index: 3, count: 2 };
function translate(formula, options = {}) {
  return rewriteStructuralFormula(formula, { ...base, ...options });
}
function changed(formula, expected, options) {
  assert.deepEqual(translate(formula, options), { ok: true, formula: expected });
}
function refused(formula, options) {
  const result = translate(formula, options);
  assert.equal(result.ok, false, formula);
  assert.equal(typeof result.reason, 'string');
  assert.ok(result.reason.length > 0);
}

test('insertion shifts references regardless of absolute markers and expands ranges internally', () => {
  changed('=SUM(A1,$B$3,C$4,$D2,A2:A4,A3:A3)', '=SUM(A1,$B$5,C$6,$D2,A2:A6,A5:A5)');
  changed('SUM(A1:B2)', 'SUM(A1:B2)');
  changed('A4:A2', 'A6:A2');
  changed('aBc3+$aBc$1', 'aBc5+$aBc$1');
});

test('deletion shrinks, moves or invalidates ranges and individual cells', () => {
  const options = { action: 'delete', index: 3, count: 2 };
  changed('A1+A3+$B$4+C5+A2:A6+A3:A4+A3:A6+A1:A4+A6:A3', 'A1+#REF!+#REF!+C3+A2:A4+#REF!+A3:A4+A1:A2+A4:A3', options);
  changed('A2:B3', 'A2:B2', options);
});

test('column and whole row or column ranges translate on the relevant axis', () => {
  changed('SUM($A:C,2:4,$2:$3,A1:C5)', 'SUM($A:C,2:6,$2:$5,A1:C7)');
  changed('SUM($A:C,2:4,$B$1:D5)', 'SUM($A:E,2:4,$D$1:F5)', { axis: 'column', index: 2 });
  changed('B:D', 'B:B', { axis: 'column', action: 'delete', index: 2, count: 2 });
});

test('sheet qualification, quoted apostrophes and host sheet determine references', () => {
  changed("='Data'!A3+Other!A3+A3+'O''Brien'!$A$3", "='Data'!A5+Other!A3+A5+'O''Brien'!$A$3");
  changed("='O''Brien'!A3+DATA!A3+A3", "='O''Brien'!A3+DATA!A5+A3", { formulaSheet: 'Other' });
  changed("='O''Brien'!A3", "='O''Brien'!A5", { sheet: "O'Brien" });
  changed('Data!A3:Data!B4', 'Data!A5:Data!B6');
});

test('strings, escaped quotes, errors, numbers, function identifiers and names stay intact', () => {
  changed('=IF(A3="A3 ""Data!B4""",LOG10(1E3)+My_Name1,#N/A)', '=IF(A5="A3 ""Data!B4""",LOG10(1E3)+My_Name1,#N/A)');
  changed('=SUM(#REF!,A3)', '=SUM(#REF!,A5)');
});

test('a call is not mistaken for intersection and ambiguity is refused', () => {
  // `LOG10` is both a function and a valid A1 address; the immediate `(` wins.
  changed('=LOG10(A3)', '=LOG10(A5)');
  changed('=SUM (A3:A4)', '=SUM (A5:A6)');
  // A cell-like name with a space before `(` is either intersection or a spaced
  // call; translating either reading could corrupt the formula, so refuse.
  refused('=A3 (A1:A4)');
  refused("='Data'!A3 (A1:A4)");
  refused('INDIRECT (A3)');
});

test('workbook defined names require sheet context for unqualified references', () => {
  changed("'Data'!$A$3:$B$4", "'Data'!$A$5:$B$6", { formulaSheet: null });
  changed('MyName+Other!A3', 'MyName+Other!A3', { formulaSheet: null });
  refused('A3', { formulaSheet: null });
});

test('bounds and invalid operation parameters fail closed', () => {
  refused('A1048576', { index: 1, count: 1 });
  refused('XFD1', { axis: 'column', index: 1, count: 1 });
  refused('A1', { index: 0 });
  refused('A1', { index: 1048576, count: 2 });
  refused('A1', { count: 0 });
  refused('A1', { axis: 'sheet' });
});

test('uncertain references and dynamic syntax fail closed without partial results', () => {
  for (const formula of [
    'Sheet1:Sheet3!A1', "'Sheet1:Sheet3'!A1", '[Book.xlsx]Data!A1',
    "'[Book.xlsx]Data'!A1", 'Table1[Amount]', '@A1', 'A1#',
    'INDIRECT("A3")', 'OFFSET(A1,1,0)', 'R1C1', 'A1:Other!B3',
    'A1:', '"unterminated', "'unterminated!A1", 'A1 + {1,2}',
  ]) refused(formula);
});
