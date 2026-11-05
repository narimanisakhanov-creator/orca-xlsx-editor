// Surgical xlsx writer: edits cell XML in place inside the package's entries.
//
// The caller owns zip reading/writing; this module only turns `ZipEntry[]` into a
// new `ZipEntry[]` where the parts it actually touched carry fresh bytes and every
// other entry is the very same object it was handed.

import {
  escapeXml,
  findElement,
  findElements,
  getAttr,
  removeAttr,
  replaceElements,
  rootName,
  rootPrefix,
  setAttr,
  topLevelElements,
} from './xml.js';

/**
 * @typedef {{ kind: 'number', value: number }
 *   | { kind: 'date', value: number }
 *   | { kind: 'text', value: string }
 *   | { kind: 'boolean', value: boolean }
 *   | { kind: 'error', value: string }
 *   | { kind: 'formula', formula: string, cached?: number | string | boolean | null }
 *   | { kind: 'blank' }} CellValue
 * @typedef {{ sheet: string, ref: string, value: CellValue, styleFrom?: string }} CellEdit
 */

const TEXT_ENCODER = new TextEncoder();

class Refusal extends Error {}

function refuse(reason) {
  return { status: 'refused', reason };
}

// ---------------------------------------------------------------------------
// Numbers, references, paths

function colToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const v = ch.charCodeAt(0) - 64;
    if (v < 1 || v > 26) return 0;
    n = n * 26 + v;
  }
  return n;
}

function indexToCol(n) {
  let s = '';
  let rest = n;
  while (rest > 0) {
    const r = (rest - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    rest = Math.floor((rest - 1) / 26);
  }
  return s;
}

const REF_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;
const MAX_ROW = 1048576;
const MAX_COL = 16384;

function parseRef(ref) {
  if (typeof ref !== 'string') return null;
  const m = REF_RE.exec(ref.trim());
  if (!m) return null;
  const col = colToIndex(m[1]);
  const row = Number(m[2]);
  if (!col || col > MAX_COL || !row || row > MAX_ROW) return null;
  return { row, col, ref: `${indexToCol(col)}${row}` };
}

function parseRange(text) {
  if (typeof text !== 'string') return null;
  const parts = text.trim().split(':');
  if (parts.length === 1) {
    const one = parseRef(parts[0]);
    return one ? { r1: one.row, c1: one.col, r2: one.row, c2: one.col } : null;
  }
  if (parts.length !== 2) return null;
  const a = parseRef(parts[0]);
  const b = parseRef(parts[1]);
  if (!a || !b) return null;
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
  };
}

/** A1-style text for a box; a single cell is written as `B3`, not `B3:B3`. */
function rangeText(box) {
  if (box.r1 === box.r2 && box.c1 === box.c2) return `${indexToCol(box.c1)}${box.r1}`;
  return `${indexToCol(box.c1)}${box.r1}:${indexToCol(box.c2)}${box.r2}`;
}

/** Excel stores `<v>` as a double literal; keep it plain decimal where it fits. */
function numberText(n) {
  if (!Number.isFinite(n)) throw new Refusal(`number ${n} cannot be written to a cell`);
  if (Object.is(n, -0)) return '0';
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const plain = String(n);
  if (!plain.includes('e')) return plain;
  return n.toExponential().replace('e', 'E');
}

function dirOf(path) {
  const i = path.lastIndexOf('/');
  return i < 0 ? '' : path.slice(0, i);
}

function baseOf(path) {
  const i = path.lastIndexOf('/');
  return i < 0 ? path : path.slice(i + 1);
}

function resolveTarget(base, target) {
  if (!target) return null;
  let raw = String(target).split('#')[0].replace(/\\/g, '/').trim();
  if (!raw) return null;
  if (!raw.startsWith('/')) raw = base ? `${base}/${raw}` : raw;
  const out = [];
  for (const seg of raw.split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') {
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join('/');
}

function relsPathFor(partPath) {
  const dir = dirOf(partPath);
  const rels = `_rels/${baseOf(partPath)}.rels`;
  return dir ? `${dir}/${rels}` : rels;
}

function findEntry(entries, path) {
  if (!path) return null;
  const wanted = String(path).toLowerCase();
  let loose = null;
  for (const entry of entries) {
    if (!entry || typeof entry.name !== 'string') continue;
    if (entry.name === path) return entry;
    if (entry.name.toLowerCase() === wanted) loose = entry;
  }
  return loose;
}

// ---------------------------------------------------------------------------
// Part text

function decodeXml(entry) {
  const bytes = entry.data;
  if (!(bytes instanceof Uint8Array)) throw new Refusal(`part "${entry.name}" has no bytes to rewrite`);
  if (
    bytes.length >= 2 &&
    ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))
  ) {
    throw new Refusal(`part "${entry.name}" is UTF-16; this writer only rewrites UTF-8 parts`);
  }
  const bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  const body = bom ? bytes.subarray(3) : bytes;
  let text;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(body);
  } catch {
    throw new Refusal(`part "${entry.name}" is not valid UTF-8`);
  }
  return { text, bom };
}

function encodeXml(text, bom) {
  const body = TEXT_ENCODER.encode(text);
  if (!bom) return body;
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf], 0);
  out.set(body, 3);
  return out;
}

// ---------------------------------------------------------------------------
// Package model

/**
 * Read the workbook skeleton from the entries: sheet names, their rel ids and the
 * worksheet part each one points at. Returns null when the package is not a workbook;
 * pass `diag` to learn why (the writer surfaces that as its refusal reason).
 * @returns {null | { entries: object[], files: Map<string, object>, book: object }}
 */
export function xlsxPackage(zipEntries, diag = null) {
  const fail = (reason) => {
    if (diag) diag.reason = reason;
    return null;
  };
  if (!Array.isArray(zipEntries)) return fail('zipEntries must be an array');
  try {
    const files = new Map();
    for (const entry of zipEntries) {
      if (entry && typeof entry.name === 'string') files.set(entry.name, entry);
    }
    const lookup = (path) => (path == null ? null : files.get(path) ?? findEntry(zipEntries, path));

    let bookPath = null;
    const rootRels = lookup('_rels/.rels');
    if (rootRels) {
      const { text } = decodeXml(rootRels);
      if (rootName(text) !== 'Relationships') return fail('package relationships part is not a Relationships document');
      for (const rel of findElements(text, 'Relationship')) {
        const type = getAttr(rel.attrs, 'Type') ?? '';
        if (!type.endsWith('/officeDocument')) continue;
        bookPath = resolveTarget('', getAttr(rel.attrs, 'Target') ?? '');
        break;
      }
    }
    if (!bookPath) bookPath = 'xl/workbook.xml';

    const bookEntry = lookup(bookPath);
    if (!bookEntry) return fail(`workbook part "${bookPath}" is missing from the package`);
    const book = decodeXml(bookEntry);
    if (rootName(book.text) !== 'workbook') return fail('workbook part is not a workbook document');

    const relsPath = relsPathFor(bookEntry.name);
    const relsEntry = lookup(relsPath);
    const rels = new Map();
    if (relsEntry) {
      const relText = decodeXml(relsEntry).text;
      if (rootName(relText) !== 'Relationships') return fail('workbook relationships part is not a Relationships document');
      for (const rel of findElements(relText, 'Relationship')) {
        const id = getAttr(rel.attrs, 'Id');
        if (!id) continue;
        rels.set(id, {
          id,
          type: getAttr(rel.attrs, 'Type') ?? '',
          target: getAttr(rel.attrs, 'Target') ?? '',
        });
      }
    }

    const bookDir = dirOf(bookEntry.name);
    const sheets = [];
    for (const el of findElements(book.text, 'sheet')) {
      const name = getAttr(el.attrs, 'name');
      if (name == null) continue;
      const relId = getAttr(el.attrs, 'r:id');
      const rel = relId ? rels.get(relId) ?? null : null;
      const part = rel ? resolveTarget(bookDir, rel.target) : null;
      sheets.push({
        name,
        sheetId: getAttr(el.attrs, 'sheetId'),
        state: getAttr(el.attrs, 'state'),
        relId,
        relType: rel ? rel.type : null,
        part,
        entry: part ? lookup(part) : null,
      });
    }

    return {
      entries: zipEntries.slice(),
      files,
      book: {
        path: bookEntry.name,
        entry: bookEntry,
        text: book.text,
        bom: book.bom,
        relsPath: relsEntry ? relsEntry.name : null,
        relsEntry,
        sheets,
      },
    };
  } catch (err) {
    if (err instanceof Refusal) return fail(err.message);
    throw err;
  }
}

// ---------------------------------------------------------------------------
// Cell payloads

// Excel's error literals, as stored in `t="e"` cells. A cell with any other text under that type
// is something no reader can classify, so it is refused instead of written.
const EXCEL_ERRORS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
]);

// The last serial Excel's date format carries: 9999-12-31 is 2958465, so 2958466 is already out.
// The bound applies to both date systems; which epoch serial 0 means is the caller's business.
const MAX_DATE_SERIAL = 2958466;

/** Excel truncates a cell's text at 32,767 UTF-16 units; longer text makes the part unreadable. */
const MAX_CELL_TEXT = 32767;

// `_xHHHH_` is Excel's own character escape, so a literal occurrence in user text must hide its
// underscore as `_x005F_`; otherwise Excel decodes `_x0041_` to `A`.
const LITERAL_X_ESCAPE = /_(x[0-9A-Fa-f]{4})_/g;
// Controls XML 1.0 cannot carry, plus CR: a literal CR is normalised away by the XML parser
// before Excel sees it, so it has to travel as `_x000D_`. Tab and LF are legal as-is.
const NON_TEXT_CONTROL = /\r\n?|[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

/** SpreadsheetML text escaping: `_xHHHH_` literals, C0 controls and CR, then XML entities. */
function escapeXlsxText(value) {
  let text = String(value).replace(LITERAL_X_ESCAPE, '_x005F_$1_');
  text = text.replace(NON_TEXT_CONTROL, (ch) => {
    if (ch === '\r\n') return '_x000D_\n';
    if (ch === '\r') return '_x000D_';
    return `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`;
  });
  return escapeXml(text);
}

/** Check Excel's per-cell text bound and return the string to write. */
function boundedCellText(value, where) {
  const text = value == null ? '' : String(value);
  if (text.length > MAX_CELL_TEXT) {
    throw new Refusal(
      `${where}: text is ${text.length} characters, over Excel's ${MAX_CELL_TEXT}-character cell limit`,
    );
  }
  return text;
}

function buildPayload(edit, prefix = '') {
  const p = prefix;
  const where = `cell ${edit.ref} on sheet "${edit.sheet}"`;
  const value = edit.value;
  switch (value.kind) {
    case 'number': {
      if (typeof value.value !== 'number' || !Number.isFinite(value.value)) {
        throw new Refusal(`${where}: number value is not a finite number`);
      }
      return { t: null, inner: `<${p}v>${numberText(value.value)}</${p}v>` };
    }
    case 'date': {
      // A date is just a number with a date format: the caller converts `Date` to a serial with
      // the workbook's own epoch, this module only writes the number and must not emit `t`.
      const serial = value.value;
      if (typeof serial !== 'number' || !Number.isFinite(serial)) {
        throw new Refusal(`${where}: date value must be a finite Excel serial number`);
      }
      if (serial < 0 || serial >= MAX_DATE_SERIAL) {
        throw new Refusal(`${where}: date serial ${numberText(serial)} is outside the representable range`);
      }
      return { t: null, inner: `<${p}v>${numberText(serial)}</${p}v>` };
    }
    case 'text': {
      const text = boundedCellText(value.value, where);
      return {
        t: 'inlineStr',
        inner: `<${p}is><${p}t xml:space="preserve">${escapeXlsxText(text)}</${p}t></${p}is>`,
      };
    }
    case 'boolean':
      return { t: 'b', inner: `<${p}v>${value.value ? 1 : 0}</${p}v>` };
    case 'error': {
      const code = value.value;
      if (typeof code !== 'string' || !EXCEL_ERRORS.has(code)) {
        throw new Refusal(`${where}: ${JSON.stringify(code)} is not an Excel error code`);
      }
      return { t: 'e', inner: `<${p}v>${escapeXml(code)}</${p}v>` };
    }
    case 'formula': {
      const formula = typeof value.formula === 'string' ? value.formula.trim() : '';
      if (!formula) throw new Refusal(`${where}: formula is empty`);
      const hasCached = value.cached !== undefined && value.cached !== null;
      if (!hasCached) return { t: null, inner: `<${p}f>${escapeXml(formula)}</${p}f>` };
      const cached = value.cached;
      if (typeof cached === 'number') {
        if (!Number.isFinite(cached)) throw new Refusal(`${where}: cached value is not finite`);
        return { t: null, inner: `<${p}f>${escapeXml(formula)}</${p}f><${p}v>${numberText(cached)}</${p}v>` };
      }
      if (typeof cached === 'boolean') {
        return { t: 'b', inner: `<${p}f>${escapeXml(formula)}</${p}f><${p}v>${cached ? 1 : 0}</${p}v>` };
      }
      const result = boundedCellText(cached, `${where} (cached formula result)`);
      return {
        t: 'str',
        inner: `<${p}f>${escapeXml(formula)}</${p}f><${p}v>${escapeXlsxText(result)}</${p}v>`,
      };
    }
    case 'blank':
      return { t: null, inner: '' };
    default:
      throw new Refusal(`${where}: unsupported value kind ${JSON.stringify(value.kind)}`);
  }
}

// ---------------------------------------------------------------------------
// sheetData surgery
//
// Edits are applied row by row. A row's text is small, so hunting cells inside it is
// cheap even on a 100k-row sheet, and each row is spliced exactly once no matter how
// many of its cells change.

/** Rows of a worksheet, in document order, with resolved row numbers. */
function rowsIn(xml, sheetData) {
  const out = [];
  let implied = 0;
  for (const el of findElements(xml, 'row')) {
    if (el.start < sheetData.innerStart || el.end > sheetData.innerEnd) continue;
    const num = Number(getAttr(el.attrs, 'r'));
    const row = Number.isInteger(num) && num > 0 ? num : implied + 1;
    implied = row;
    out.push({ el, num: row });
  }
  return out;
}

/** Cells inside a row's inner text, in document order, with resolved column numbers. */
function cellsOf(rowText) {
  const out = [];
  let implied = 0;
  for (const el of findElements(rowText, 'c')) {
    const parsed = parseRef(getAttr(el.attrs, 'r') ?? '');
    const col = parsed ? parsed.col : implied + 1;
    implied = col;
    out.push({ el, col });
  }
  return out;
}

/** A cell-level `extLst` carries metadata Excel trusts; rewriting the cell must keep it. */
function preservedExtLst(cellText, cell) {
  if (cell.el.selfClosing) return '';
  const inner = cellText.slice(cell.el.start, cell.el.end);
  const ext = findElement(inner, 'extLst');
  return ext ? inner.slice(ext.start, ext.end) : '';
}

/** Style of an existing cell, for an explicit `styleFrom` hint. Never invented. */
function readCellStyle(sheetXml, refText) {
  const target = parseRef(refText);
  if (!target) return null;
  const sheetData = topChild(sheetXml, 'sheetData');
  if (!sheetData) return null;
  const row = rowsIn(sheetXml, sheetData).find((candidate) => candidate.num === target.row);
  if (!row) return null;
  const rowText = sheetXml.slice(row.el.innerStart, row.el.innerEnd);
  const cell = cellsOf(rowText).find((candidate) => candidate.col === target.col);
  return cell ? getAttr(cell.el.attrs, 's') : null;
}

function isBlankPayload(payload) {
  return payload.inner === '';
}

/** Full `<c>` text for a cell that does not exist yet; the root's prefix is kept. */
function newCellText(edit, payload, styleXml, prefix) {
  let style = null;
  if (typeof edit.styleFrom === 'string') style = readCellStyle(styleXml, edit.styleFrom);
  const sAttr = style == null ? '' : ` s="${escapeXml(style)}"`;
  if (isBlankPayload(payload)) return `<${prefix}c r="${edit.ref}"${sAttr}/>`;
  const tAttr = payload.t ? ` t="${payload.t}"` : '';
  return `<${prefix}c r="${edit.ref}"${sAttr}${tAttr}>${payload.inner}</${prefix}c>`;
}

/** New text for an existing cell: style kept, value replaced, extLst preserved. */
function existingCellText(cellText, cell, payload) {
  // The cell's own tag name carries any namespace prefix the part was written with.
  const tag = cell.el.name;
  const attrs = removeAttr(cell.el.attrs, 't');
  if (isBlankPayload(payload)) {
    // Blank keeps `s` so Excel still treats the slot as formatted; a cell with nothing
    // but its reference left carries no information, so it goes away entirely.
    if (removeAttr(attrs, 'r').trim() === '') return '';
    return `<${tag}${attrs}/>`;
  }
  const ext = preservedExtLst(cellText, cell);
  const tAttr = payload.t ? ` t="${payload.t}"` : '';
  return `<${tag}${attrs}${tAttr}>${payload.inner}${ext}</${tag}>`;
}

/** Apply one row's edits to that row's inner text. */
function rewriteRowText(rowText, cells, edits, styleXml, prefix) {
  // A repeated edit of one cell resolves to the last one, the way a real overwrite does.
  const latest = new Map();
  for (const edit of edits) latest.set(edit.col, edit);

  const present = new Set(cells.map((cell) => cell.col));
  // Every new cell of the row is collected first and written as one ascending run. Planning
  // them one at a time gave two cells in the same gap the same insertion offset, and splicing
  // back to front then reversed their columns — a `<c>` order Excel refuses to open.
  const additions = [...latest.values()]
    .filter((edit) => !present.has(edit.col) && !isBlankPayload(edit.payload))
    .sort((a, b) => a.col - b.col);

  let out = '';
  let cursor = 0;
  let next = 0;
  for (const cell of cells) {
    while (next < additions.length && additions[next].col < cell.col) {
      out += newCellText(additions[next], additions[next].payload, styleXml, prefix);
      next += 1;
    }
    out += rowText.slice(cursor, cell.el.start);
    const edit = latest.get(cell.col);
    out += edit
      ? existingCellText(rowText, cell, edit.payload)
      : rowText.slice(cell.el.start, cell.el.end);
    cursor = cell.el.end;
  }
  // CT_Row is `c*` then `extLst?`: cells past the last existing one go before the row's
  // metadata, never after it.
  while (next < additions.length) {
    out += newCellText(additions[next], additions[next].payload, styleXml, prefix);
    next += 1;
  }
  out += rowText.slice(cursor);
  return out;
}

/**
 * Apply every edit of one sheet to its worksheet XML.
 * @returns {{ xml: string, applied: object[] }} `applied` excludes blanks that hit
 *   cells which never existed, so the dimension is not widened for a no-op.
 */
function writeRows(xml, sheetData, edits) {
  const prefix = rootPrefix(xml);
  const rows = rowsIn(xml, sheetData);
  // Payloads are built first so a bad value refuses before anything is spliced.
  const prepared = edits.map((edit) => ({ ...edit, payload: buildPayload(edit, prefix) }));
  const byRow = new Map();
  for (const edit of prepared) {
    let list = byRow.get(edit.row);
    if (!list) byRow.set(edit.row, (list = []));
    const previous = list.findIndex((candidate) => candidate.col === edit.col);
    if (previous >= 0) list[previous] = edit;
    else list.push(edit);
  }
  // Descending rows: inserting at a higher row never shifts the offsets of a lower one.
  const order = [...byRow.keys()].sort((a, b) => b - a);

  const applied = [];
  let out = xml;
  for (const rowNum of order) {
    const rowEdits = byRow.get(rowNum).sort((a, b) => a.col - b.col);
    const row = rows.find((candidate) => candidate.num === rowNum);

    if (!row || row.el.selfClosing) {
      const writable = rowEdits.filter((edit) => !isBlankPayload(edit.payload));
      if (writable.length === 0) continue;
      const cells = writable
        .map((edit) => newCellText(edit, edit.payload, xml, prefix))
        .join('');
      if (row) {
        // Rebuild the start tag from its attribute text: `<row r="1" />` must reopen as
        // `<row r="1">`, and string surgery on `/>` leaves stray whitespace or `</row>`.
        const tag = row.el.name;
        out =
          out.slice(0, row.el.start) +
          `<${tag}${row.el.attrs}>${cells}</${tag}>` +
          out.slice(row.el.end);
      } else {
        const anchor = rows.find((candidate) => candidate.num > rowNum);
        const at = anchor ? anchor.el.start : sheetData.innerEnd;
        out =
          out.slice(0, at) +
          `<${prefix}row r="${rowNum}">${cells}</${prefix}row>` +
          out.slice(at);
      }
      applied.push(...writable);
      continue;
    }

    const rowText = out.slice(row.el.innerStart, row.el.innerEnd);
    const cells = cellsOf(rowText);
    const next = rewriteRowText(rowText, cells, rowEdits, xml, prefix);
    out = out.slice(0, row.el.innerStart) + next + out.slice(row.el.innerEnd);
    for (const edit of rowEdits) {
      const existed = cells.some((cell) => cell.col === edit.col);
      if (existed || !isBlankPayload(edit.payload)) applied.push(edit);
    }
  }
  return { xml: out, applied };
}

const AFTER_DIMENSION = [
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
];

const AFTER_CALC_PR = [
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
];

/** Insert a snippet as a direct child of the root, before the first of `before`. */
function insertTopLevel(xml, snippet, before) {
  const root = findElement(xml, rootName(xml) ?? '');
  if (!root || root.selfClosing) throw new Refusal('document root cannot host new elements');
  const anchor = topLevelElements(xml).find((child) => before.includes(child.local));
  const at = anchor ? anchor.start : root.innerEnd;
  return xml.slice(0, at) + snippet + xml.slice(at);
}

function topChild(xml, name) {
  return topLevelElements(xml).find((child) => child.local === name) ?? null;
}

/** Every calcChain part the package actually carries: via its relationship, plus the default path. */
function calcChainEntries(pkg) {
  const found = new Map();
  const bookDir = dirOf(pkg.book.path);
  if (pkg.book.relsEntry) {
    const { text } = decodeXml(pkg.book.relsEntry);
    for (const rel of findElements(text, 'Relationship')) {
      const type = getAttr(rel.attrs, 'Type') ?? '';
      if (!type.endsWith('/calcChain')) continue;
      const path = resolveTarget(bookDir, getAttr(rel.attrs, 'Target') ?? '');
      const entry = path ? findEntry(pkg.entries, path) : null;
      if (entry) found.set(entry.name, entry);
    }
  }
  const fallback = findEntry(pkg.entries, `${bookDir || 'xl'}/calcChain.xml`);
  if (fallback) found.set(fallback.name, fallback);
  return [...found.values()];
}

function boxOfCells(xml) {
  let box = null;
  for (const el of findElements(xml, 'c')) {
    const parsed = parseRef(getAttr(el.attrs, 'r') ?? '');
    if (!parsed) continue;
    if (!box) box = { r1: parsed.row, c1: parsed.col, r2: parsed.row, c2: parsed.col };
    else {
      box.r1 = Math.min(box.r1, parsed.row);
      box.c1 = Math.min(box.c1, parsed.col);
      box.r2 = Math.max(box.r2, parsed.row);
      box.c2 = Math.max(box.c2, parsed.col);
    }
  }
  return box;
}

function syncDimension(xml, edits, sheetName) {
  const dim = topChild(xml, 'dimension');
  const declared = dim ? (getAttr(dim.attrs, 'ref') ?? '').trim() : '';
  // A `<dimension/>` without a reference is legal-ish and simply means "not declared";
  // only a reference that is present but unparsable is a reason to refuse.
  let box = declared ? parseRange(declared) : null;
  if (declared && !box) {
    throw new Refusal(`sheet "${sheetName}" has an unreadable dimension reference ${JSON.stringify(declared)}`);
  }
  if (!box) box = boxOfCells(xml);
  for (const edit of edits) {
    if (!box) box = { r1: edit.row, c1: edit.col, r2: edit.row, c2: edit.col };
    else {
      box.r1 = Math.min(box.r1, edit.row);
      box.c1 = Math.min(box.c1, edit.col);
      box.r2 = Math.max(box.r2, edit.row);
      box.c2 = Math.max(box.c2, edit.col);
    }
  }
  if (!box) return xml;
  const ref = rangeText(box);
  if (dim) {
    if (declared === ref) return xml;
    const attrs = setAttr(dim.attrs, 'ref', ref);
    const tag = dim.name;
    const replacement = dim.selfClosing
      ? `<${tag}${attrs}/>`
      : `<${tag}${attrs}>${xml.slice(dim.innerStart, dim.innerEnd)}</${tag}>`;
    return xml.slice(0, dim.start) + replacement + xml.slice(dim.end);
  }
  return insertTopLevel(xml, `<${rootPrefix(xml)}dimension ref="${ref}"/>`, AFTER_DIMENSION);
}

/**
 * Refuse the edits Excel would silently undo.
 *
 * A shared formula master owns its range: every slave's `<f t="shared" si>` points at the
 * master's `<f t="shared" ref>`, so replacing the master with an ordinary value leaves the
 * slaves aiming at a formula that no longer exists and Excel rejects the whole book. A slave
 * carries no `ref`; replacing its `<f>` with a value simply makes that cell independent.
 *
 * An array formula is recalculated from its anchor, so a value written into any other cell of
 * its range disappears on the next calculation.
 */
function guardEditTargets(xml, sheetName, edits, notes) {
  const arrays = findElements(xml, 'f')
    .filter((el) => (getAttr(el.attrs, 't') ?? '') === 'array')
    .map((el) => parseRange(getAttr(el.attrs, 'ref') ?? ''))
    .filter(Boolean);

  const sheetData = topChild(xml, 'sheetData');
  const rows = sheetData ? new Map(rowsIn(xml, sheetData).map((row) => [row.num, row])) : new Map();

  for (const edit of edits) {
    for (const box of arrays) {
      const inside =
        edit.row >= box.r1 && edit.row <= box.r2 && edit.col >= box.c1 && edit.col <= box.c2;
      if (!inside) continue;
      if (edit.row === box.r1 && edit.col === box.c1) continue;
      throw new Refusal(
        `cell ${edit.ref} on sheet "${sheetName}" is inside array formula range ${rangeText(box)}; ` +
          `only ${indexToCol(box.c1)}${box.r1} is writable`,
      );
    }
    const row = rows.get(edit.row);
    if (!row || row.el.selfClosing) continue;
    const rowText = xml.slice(row.el.innerStart, row.el.innerEnd);
    const cell = cellsOf(rowText).find((candidate) => candidate.col === edit.col);
    if (!cell || cell.el.selfClosing) continue;
    const formula = findElement(rowText.slice(cell.el.innerStart, cell.el.innerEnd), 'f');
    if (!formula || (getAttr(formula.attrs, 't') ?? '') !== 'shared') continue;
    if (getAttr(formula.attrs, 'ref') == null) continue;
    throw new Refusal(
      `cell ${edit.ref} on sheet "${sheetName}" is a shared formula master at ${edit.ref}; ` +
        'edit a slave or the whole range',
    );
  }

  // `cm="1"` marks a dynamic-array anchor, but nothing in the sheet says how far the spill
  // reaches — the spilled cells usually are not in the file at all. A guessed range would
  // refuse legitimate edits, so only the proven `t="array" ref` case is enforced and this
  // limitation is reported instead.
  if (findElements(xml, 'c').some((el) => getAttr(el.attrs, 'cm') != null)) {
    notes.push(
      `sheet "${sheetName}": cm="1" spill markers are present; spill ranges are not tracked, ` +
        'only t="array" ranges are protected',
    );
  }
}

function applySheetEdits(text, sheetName, edits, notes) {
  const protection = topChild(text, 'sheetProtection');
  if (protection) {
    const flag = (getAttr(protection.attrs, 'sheet') ?? '1').toLowerCase();
    if (flag !== '0' && flag !== 'false') {
      throw new Refusal(`sheet "${sheetName}" is protected (sheetProtection); refusing to edit cells`);
    }
  }

  const merges = [];
  const mergeCells = topChild(text, 'mergeCells');
  if (mergeCells) {
    for (const el of findElements(text, 'mergeCell')) {
      if (el.start < mergeCells.innerStart || el.end > mergeCells.innerEnd) continue;
      const parsed = parseRange(getAttr(el.attrs, 'ref') ?? '');
      if (!parsed) throw new Refusal(`sheet "${sheetName}" has an unreadable merged range`);
      merges.push(parsed);
    }
  }
  for (const edit of edits) {
    for (const merge of merges) {
      const inside =
        edit.row >= merge.r1 && edit.row <= merge.r2 && edit.col >= merge.c1 && edit.col <= merge.c2;
      if (!inside) continue;
      if (edit.row === merge.r1 && edit.col === merge.c1) continue;
      throw new Refusal(
        `cell ${edit.ref} on sheet "${sheetName}" is inside merged range ${rangeText(merge)}; ` +
          `only ${indexToCol(merge.c1)}${merge.r1} is writable`,
      );
    }
  }

  let xml = text;
  const sd = topChild(xml, 'sheetData');
  if (!sd) throw new Refusal(`sheet "${sheetName}" has no sheetData element`);
  if (sd.selfClosing) {
    // `<sheetData />` becomes `<sheetData>…</sheetData>` with its attributes intact.
    xml = `${xml.slice(0, sd.start)}<sheetData${sd.attrs}></sheetData>${xml.slice(sd.end)}`;
  }

  guardEditTargets(xml, sheetName, edits, notes);
  const written = writeRows(xml, topChild(xml, 'sheetData'), edits);
  xml = syncDimension(written.xml, written.applied, sheetName);
  notes.push(`sheet "${sheetName}": wrote ${written.applied.length} of ${edits.length} requested cell(s)`);
  return xml;
}

// ---------------------------------------------------------------------------
// Workbook-level recalculation bookkeeping

/** Delete matching elements; `replaceElements` already handles the offset arithmetic. */
function removeElements(xml, name, predicate) {
  return replaceElements(xml, name, (el) => (predicate(el) ? '' : null));
}

function setCalcOnLoad(xml) {
  const calcPr = topChild(xml, 'calcPr');
  if (calcPr) {
    const current = getAttr(calcPr.attrs, 'fullCalcOnLoad');
    if (current === '1' || current === 'true') return xml;
    const attrs = setAttr(calcPr.attrs, 'fullCalcOnLoad', '1');
    const tag = calcPr.name;
    const replacement = calcPr.selfClosing
      ? `<${tag}${attrs}/>`
      : `<${tag}${attrs}>${xml.slice(calcPr.innerStart, calcPr.innerEnd)}</${tag}>`;
    return xml.slice(0, calcPr.start) + replacement + xml.slice(calcPr.end);
  }
  const prefix = rootPrefix(xml);
  return insertTopLevel(xml, `<${prefix}calcPr fullCalcOnLoad="1"/>`, AFTER_CALC_PR);
}

// ---------------------------------------------------------------------------
// Public entry point

/**
 * Apply cell edits to the package entries.
 * @param {object[]} zipEntries
 * @param {CellEdit[]} edits
 * @returns {{ status: 'ok', entries: object[], notes: string[] } | { status: 'refused', reason: string }}
 */
export function applyEdits(zipEntries, edits) {
  try {
    if (!Array.isArray(zipEntries)) throw new Refusal('zipEntries must be an array');
    if (!Array.isArray(edits)) throw new Refusal('edits must be an array');
    const list = edits.map((edit, index) => {
      if (!edit || typeof edit !== 'object') throw new Refusal(`edit #${index} is not an object`);
      if (typeof edit.sheet !== 'string' || edit.sheet === '') {
        throw new Refusal(`edit #${index} has no sheet name`);
      }
      const parsed = parseRef(edit.ref);
      if (!parsed) throw new Refusal(`edit #${index} has an invalid cell reference ${JSON.stringify(edit.ref)}`);
      if (!edit.value || typeof edit.value !== 'object' || typeof edit.value.kind !== 'string') {
        throw new Refusal(`edit #${index} has no value`);
      }
      return { ...edit, ref: parsed.ref, row: parsed.row, col: parsed.col };
    });
    if (list.length === 0) return { status: 'ok', entries: zipEntries.slice(), notes: ['no cell edits'] };

    const diag = { reason: '' };
    const pkg = xlsxPackage(zipEntries, diag);
    if (!pkg) throw new Refusal(diag.reason || 'workbook part is missing or unreadable');

    const notes = [];
    /** @type {Map<string, { before: string, after: string, bom: boolean }>} */
    const pending = new Map();
    const markChanged = (entry, before, after, bom) => {
      if (after === before) return;
      pending.set(entry.name, { before, after, bom });
    };

    // Group by sheet, case-insensitively: Excel sheet names are case-insensitive.
    const groups = new Map();
    for (const edit of list) {
      const key = edit.sheet.toLowerCase();
      let group = groups.get(key);
      if (!group) groups.set(key, (group = { name: edit.sheet, edits: [] }));
      group.edits.push(edit);
    }

    for (const [key, group] of groups) {
      const sheet = pkg.book.sheets.find((candidate) => candidate.name.toLowerCase() === key);
      if (!sheet) throw new Refusal(`sheet "${group.name}" not found in the workbook`);
      if (!sheet.part) throw new Refusal(`sheet "${sheet.name}" has no worksheet relationship`);
      if (!sheet.entry) throw new Refusal(`sheet "${sheet.name}": worksheet part "${sheet.part}" is missing`);
      const { text, bom } = decodeXml(sheet.entry);
      const root = rootName(text);
      if (root !== 'worksheet') {
        throw new Refusal(`part "${sheet.entry.name}" is not a worksheet (root element ${String(root)})`);
      }
      const next = applySheetEdits(text, sheet.name, group.edits, notes);
      markChanged(sheet.entry, text, next, bom);
    }

    if (pending.size === 0) {
      // Nothing in any worksheet actually moved, so leave the workbook alone: rewriting
      // calcPr or dropping a calc chain for a no-op edit would be pointless churn.
      notes.push('no cell content changed; package left untouched');
      return { status: 'ok', entries: zipEntries.slice(), notes };
    }

    // A stale calc chain after an edit makes Excel show wrong numbers, so it goes —
    // together with the relationship and content-type override that point at it.
    const calcEntries = calcChainEntries(pkg);
    const dropped = new Set(calcEntries.map((entry) => entry.name.toLowerCase()));
    for (const entry of calcEntries) notes.push(`removed ${entry.name}`);
    if (pkg.book.relsEntry) {
      const { text, bom } = decodeXml(pkg.book.relsEntry);
      const next = removeElements(text, 'Relationship', (el) => {
        const type = getAttr(el.attrs, 'Type') ?? '';
        const target = getAttr(el.attrs, 'Target') ?? '';
        return type.endsWith('/calcChain') || /(^|\/)calcChain\.xml$/i.test(target);
      });
      if (next !== text) {
        markChanged(pkg.book.relsEntry, text, next, bom);
        notes.push(`dropped the calcChain relationship from ${pkg.book.relsEntry.name}`);
      }
    }
    const contentTypes = findEntry(pkg.entries, '[Content_Types].xml');
    if (contentTypes) {
      const { text, bom } = decodeXml(contentTypes);
      const next = removeElements(text, 'Override', (el) => {
        const part = getAttr(el.attrs, 'PartName') ?? '';
        return /(^|\/)calcChain\.xml$/i.test(part);
      });
      if (next !== text) {
        markChanged(contentTypes, text, next, bom);
        notes.push(`dropped the calcChain content-type override from ${contentTypes.name}`);
      }
    }

    const bookEntry = pkg.book.entry;
    const bookNext = setCalcOnLoad(pkg.book.text);
    if (bookNext !== pkg.book.text) {
      markChanged(bookEntry, pkg.book.text, bookNext, pkg.book.bom);
      notes.push(`set calcPr fullCalcOnLoad="1" in ${bookEntry.name}`);
    }

    const out = [];
    for (const entry of zipEntries) {
      if (dropped.has(String(entry.name).toLowerCase())) continue;
      const record = pending.get(entry.name);
      if (record && record.after !== record.before) {
        out.push({ ...entry, data: encodeXml(record.after, record.bom), changed: true });
      } else {
        out.push(entry);
      }
    }
    return { status: 'ok', entries: out, notes };
  } catch (err) {
    if (err instanceof Refusal) return refuse(err.message);
    throw err;
  }
}
