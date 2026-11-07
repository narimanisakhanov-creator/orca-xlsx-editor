// Structural edits move references, including absolute references. This is not
// the copy/fill formula translator: '$' controls copying, not structural edits.
const MAX_ROW = 1048576;
const MAX_COLUMN = 16384;
const IDENTIFIER = /^[\p{L}_\\][\p{L}\p{N}_.\\]*/u;
const NUMBER = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/;
const ERROR = /^#(?:REF!|DIV\/0!|VALUE!|N\/A|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!)/i;

function columnNumber(text) {
  let value = 0;
  for (const ch of text.toUpperCase()) value = value * 26 + ch.charCodeAt(0) - 64;
  return value;
}

function columnName(value, lowercase) {
  let text = '';
  while (value > 0) {
    value--;
    text = String.fromCharCode(65 + value % 26) + text;
    value = Math.floor(value / 26);
  }
  return lowercase ? text.toLowerCase() : text;
}

function cellParts(text) {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(text);
  if (!match) return null;
  const column = columnNumber(match[2]);
  const row = Number(match[4]);
  if (column > MAX_COLUMN || row > MAX_ROW) return null;
  return { column, row, columnText: match[2], originalColumn: column, columnMark: match[1], rowMark: match[3], lower: match[2] === match[2].toLowerCase() };
}

function rangePart(token) {
  if (token.kind === 'cell') return { kind: 'cell', ...cellParts(token.text) };
  let m = /^(\$?)([A-Za-z]{1,3})$/.exec(token.text);
  if (m && columnNumber(m[2]) <= MAX_COLUMN) return { kind: 'column', column: columnNumber(m[2]), originalColumn: columnNumber(m[2]), columnText: m[2], columnMark: m[1], lower: m[2] === m[2].toLowerCase() };
  m = /^(\$?)([1-9]\d*)$/.exec(token.text);
  if (m && Number(m[2]) <= MAX_ROW) return { kind: 'row', row: Number(m[2]), rowMark: m[1] };
  return null;
}

function render(part) {
  const col = part.column === undefined ? '' : part.columnMark + (part.column === part.originalColumn ? part.columnText : columnName(part.column, part.lower));
  const row = part.row === undefined ? '' : part.rowMark + part.row;
  return col + row;
}

// The remaining interval is contiguous after deletion; preserve endpoint order
// and each endpoint's own absolute markers.
function shiftInterval(a, b, index, count, action, limit) {
  const reverse = a > b;
  let low = Math.min(a, b);
  let high = Math.max(a, b);
  if (action === 'insert') {
    if (low >= index) low += count;
    if (high >= index) high += count;
  } else {
    const end = index + count - 1;
    if (high < index) return [a, b];
    if (low > end) { low -= count; high -= count; }
    else {
      const hasBefore = low < index;
      const hasAfter = high > end;
      if (!hasBefore && !hasAfter) return null;
      low = hasBefore ? low : index;
      high = hasAfter ? high - count : index - 1;
    }
  }
  if (low < 1 || high > limit) return false;
  return reverse ? [high, low] : [low, high];
}

function lex(formula) {
  const tokens = [];
  let pos = 0;
  while (pos < formula.length) {
    const start = pos;
    const c = formula[pos];
    if (/\s/.test(c)) {
      while (pos < formula.length && /\s/.test(formula[pos])) pos++;
      tokens.push({ kind: 'space', text: formula.slice(start, pos) });
      continue;
    }
    if (c === '"') {
      pos++;
      let closed = false;
      while (pos < formula.length) {
        if (formula[pos++] === '"') {
          if (formula[pos] === '"') pos++;
          else { closed = true; break; }
        }
      }
      if (!closed) return null;
      tokens.push({ kind: 'literal', text: formula.slice(start, pos) });
      continue;
    }
    // A sheet qualifier is a single token even when it contains doubled quotes.
    if (c === "'") {
      pos++;
      let name = '';
      let closed = false;
      while (pos < formula.length) {
        const ch = formula[pos++];
        if (ch === "'") {
          if (formula[pos] === "'") { name += "'"; pos++; }
          else { closed = true; break; }
        } else name += ch;
      }
      if (!closed || formula[pos] !== '!' || !name || /[\[\]:]/.test(name)) return null;
      pos++;
      tokens.push({ kind: 'sheet', name, text: formula.slice(start, pos) });
      continue;
    }
    const rest = formula.slice(pos);
    const qualifier = /^([\p{L}\p{N}_\\.]+)!/u.exec(rest);
    if (qualifier) {
      pos += qualifier[0].length;
      tokens.push({ kind: 'sheet', name: qualifier[1], text: qualifier[0] });
      continue;
    }
    if (c === '#') {
      const error = ERROR.exec(rest);
      if (!error) return null;
      pos += error[0].length;
      tokens.push({ kind: 'literal', text: error[0] });
      continue;
    }
    // Dollar-marked cells/whole-axis endpoints are not identifiers.
    const marked = /^\$?(?:[A-Za-z]{1,3}\$[1-9]\d*|\$[A-Za-z]{1,3}[1-9]\d*|\$[A-Za-z]{1,3}|\$[1-9]\d*)/.exec(rest);
    if (marked) {
      pos += marked[0].length;
      if (/^[\p{L}\p{N}_.\\$]/u.test(formula.slice(pos))) return null;
      const part = cellParts(marked[0]);
      tokens.push({ kind: part ? 'cell' : 'marked', text: marked[0] });
      continue;
    }
    const identifier = IDENTIFIER.exec(rest);
    if (identifier) {
      pos += identifier[0].length;
      const text = identifier[0];
      if (/^R(?:\d+|\[.*\])C(?:\d+|\[.*\])$/i.test(text)) return null;
      const tail = formula.slice(pos);
      const immediateCall = tail.startsWith('(');
      const spacedCall = /^\s+\(/.test(tail);
      const part = cellParts(text);
      // `LOG10(100)` is a call even though `LOG10` is also a valid A1 address.
      // A cell-like identifier with a space before `(` (`A3 (A1:A4)`) is either
      // the intersection operator or a spaced call: genuinely ambiguous, refuse
      // instead of silently translating the wrong reading.
      if (spacedCall && part) return null;
      const functionCall = immediateCall || spacedCall;
      if (functionCall && /(?:^|\.)(?:INDIRECT|OFFSET)$/i.test(text)) return null;
      tokens.push({ kind: functionCall ? 'name' : part ? 'cell' : 'name', text });
      continue;
    }
    const number = NUMBER.exec(rest);
    if (number) {
      pos += number[0].length;
      if (/^[\p{L}_\\]/u.test(formula.slice(pos))) return null;
      tokens.push({ kind: 'number', text: number[0] });
      continue;
    }
    if ('=+-*/^&%(),;:<>'.includes(c)) {
      tokens.push({ kind: c === ':' ? 'colon' : 'operator', text: c });
      pos++;
      continue;
    }
    return null;
  }
  return tokens;
}

/** Rewrite references for one insertion/deletion, or refuse uncertain syntax. */
export function rewriteStructuralFormula(formula, options) {
  const { sheet, formulaSheet, axis, action, index, count } = options || {};
  const limit = axis === 'row' ? MAX_ROW : MAX_COLUMN;
  if (typeof formula !== 'string' || typeof sheet !== 'string' || !sheet ||
      (formulaSheet != null && (typeof formulaSheet !== 'string' || !formulaSheet)) ||
      !['row', 'column'].includes(axis) || !['insert', 'delete'].includes(action) ||
      !Number.isInteger(index) || !Number.isInteger(count) || index < 1 || count < 1 || index + count - 1 > limit) {
    return { ok: false, reason: 'Invalid structural formula arguments' };
  }
  const tokens = lex(formula);
  if (!tokens) return { ok: false, reason: 'Unsupported or ambiguous formula syntax' };
  const fail = (reason) => ({ ok: false, reason });
  const equalSheet = (a, b) => a.toLowerCase() === b.toLowerCase();
  const nextNonSpace = (i) => { while (i < tokens.length && tokens[i].kind === 'space') i++; return i; };
  let output = '';
  for (let i = 0; i < tokens.length;) {
    let prefix = '';
    let sheetName = null;
    if (tokens[i].kind === 'sheet') {
      prefix = tokens[i].text;
      sheetName = tokens[i].name;
      i++;
      if (i >= tokens.length || tokens[i].kind === 'space' || tokens[i].kind === 'sheet') return fail('Unsupported sheet-qualified expression');
    }
    const token = tokens[i];
    if (token.kind === 'colon') return fail('Unsupported range expression');
    const colon = nextNonSpace(i + 1);
    if (colon < tokens.length && tokens[colon].kind === 'colon') {
      const first = rangePart(token);
      let j = nextNonSpace(colon + 1);
      let secondPrefix = '';
      if (tokens[j]?.kind === 'sheet') {
        if (!sheetName || !equalSheet(sheetName, tokens[j].name)) return fail('Unsupported cross-sheet range');
        secondPrefix = tokens[j++].text;
      }
      const second = tokens[j] && rangePart(tokens[j]);
      if (!first || !second || first.kind !== second.kind) return fail('Unsupported range expression');
      if (!sheetName && formulaSheet == null) return fail('Unqualified reference has no sheet context');
      const applies = equalSheet(sheetName || formulaSheet, sheet);
      if (applies && (first.kind === 'cell' || first.kind === axis)) {
        const shifted = shiftInterval(first[axis], second[axis], index, count, action, limit);
        if (shifted === false) return fail('Reference exceeds worksheet bounds');
        if (shifted === null) output += prefix + '#REF!';
        else {
          first[axis] = shifted[0]; second[axis] = shifted[1];
          output += prefix + render(first) + tokens.slice(i + 1, colon + 1).map((t) => t.text).join('') +
            tokens.slice(colon + 1, secondPrefix ? j - 1 : j).map((t) => t.text).join('') + secondPrefix + render(second);
        }
      } else output += prefix + tokens.slice(i, j + 1).map((t) => t.text).join('');
      i = j + 1;
      continue;
    }
    if (token.kind === 'marked') return fail('Invalid or incomplete reference');
    if (sheetName && token.kind !== 'cell' && token.kind !== 'name' && token.text.toUpperCase() !== '#REF!') return fail('Unsupported sheet-qualified expression');
    if (token.kind === 'cell') {
      if (!sheetName && formulaSheet == null) return fail('Unqualified reference has no sheet context');
      const part = cellParts(token.text);
      if (equalSheet(sheetName || formulaSheet, sheet)) {
        const shifted = shiftInterval(part[axis], part[axis], index, count, action, limit);
        if (shifted === false) return fail('Reference exceeds worksheet bounds');
        output += prefix + (shifted === null ? '#REF!' : (part[axis] = shifted[0], render(part)));
      } else output += prefix + token.text;
    } else output += prefix + token.text;
    i++;
  }
  return { ok: true, formula: output };
}
