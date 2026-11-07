// Prepare object part replacements without mutating the original package.
import { attrRecords, escapeXml, findElement, findElements, getAttr, localName, setAttr, topLevelElements, unescapeXml } from './xml.js';
import { rewriteStructuralFormula } from './structural-formula.js';

const OFFICE = new Set(['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships']);
const PACKAGE = new Set(['http://schemas.openxmlformats.org/package/2006/relationships', 'http://purl.oclc.org/ooxml/package/relationships']);
const SHEET = new Set(['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main']);
const DRAWING = new Set(['http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing']);
const decoder = new TextDecoder('utf-8', { fatal: true });
class Refusal extends Error {}
const refuse = reason => { throw new Refusal(reason); };

function read(entry) {
  if (!entry || !(entry.data instanceof Uint8Array)) refuse('missing object part');
  const bom = entry.data.length >= 3 && entry.data[0] === 0xef && entry.data[1] === 0xbb && entry.data[2] === 0xbf;
  try { return { text: decoder.decode(bom ? entry.data.subarray(3) : entry.data), bom }; }
  catch { refuse(`invalid UTF-8 in ${entry.name}`); }
}
function namespaces(attrs, parent = new Map()) {
  const map = new Map(parent);
  const seen = new Set();
  for (const a of attrRecords(attrs)) {
    if (a.name !== 'xmlns' && !a.name.startsWith('xmlns:')) continue;
    const prefix = a.name === 'xmlns' ? '' : a.name.slice(6);
    if (seen.has(prefix)) refuse('duplicate namespace declaration');
    seen.add(prefix); map.set(prefix, unescapeXml(a.value));
  }
  return map;
}
function nsName(el, inherited) {
  const map = namespaces(el.attrs, inherited);
  const pos = el.name.indexOf(':');
  const prefix = pos < 0 ? '' : el.name.slice(0, pos);
  return { ns: map.get(prefix), map };
}
function root(text, name, allowed) {
  const el = findElement(text, name);
  if (!el || text.slice(0, el.start).replace(/<\?xml[\s\S]*?\?>|<!--([\s\S]*?)-->/g, '').trim() || text.slice(el.end).trim()) refuse(`invalid ${name} document`);
  const info = nsName(el, new Map());
  if (!allowed.has(info.ns)) refuse(`unsupported ${name} namespace`);
  return { el, map: info.map, ns: info.ns };
}
function children(text, el) {
  return topLevelElements(text.slice(el.start, el.end)).map(child => ({ ...child, start: child.start + el.start, end: child.end + el.start, innerStart: child.innerStart + el.start, innerEnd: child.innerEnd + el.start, tagEnd: child.tagEnd + el.start }));
}
function relationshipId(el, inherited) {
  const map = namespaces(el.attrs, inherited);
  let id = null;
  for (const a of attrRecords(el.attrs)) {
    const pos = a.name.indexOf(':');
    if (pos < 1 || a.name.slice(pos + 1) !== 'id') continue;
    const ns = map.get(a.name.slice(0, pos));
    if (!ns) refuse('undeclared relationship attribute prefix');
    if (!OFFICE.has(ns)) continue;
    if (id !== null) refuse('ambiguous relationship id');
    id = unescapeXml(a.value);
  }
  if (!id) refuse('missing relationship id');
  return id;
}
function relsPath(path) {
  const slash = path.lastIndexOf('/');
  return `${slash < 0 ? '' : path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
}
function targetPath(source, target) {
  if (!target || /[?#\\]/.test(target) || /^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith('//')) refuse('unsupported relationship target');
  let decoded;
  try { decoded = decodeURIComponent(target); } catch { refuse('invalid relationship target'); }
  if (/[?#\\]/.test(decoded) || /%2f|%5c/i.test(target) || decoded.includes('\0')) refuse('unsafe relationship target');
  const segments = decoded.startsWith('/') ? [] : source.split('/').slice(0, -1);
  for (const part of decoded.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!segments.length) refuse('relationship target escapes package'); segments.pop(); }
    else segments.push(part);
  }
  if (!segments.length) refuse('empty relationship target');
  return segments.join('/');
}
function relationships(entries, source) {
  const entry = entries.get(relsPath(source));
  if (!entry) return new Map();
  const { text } = read(entry);
  const doc = root(text, 'Relationships', PACKAGE);
  const out = new Map();
  for (const el of children(text, doc.el)) {
    if (localName(el.name) !== 'Relationship' || nsName(el, doc.map).ns !== doc.ns) refuse('unsupported relationship element');
    const id = getAttr(el.attrs, 'Id'), type = getAttr(el.attrs, 'Type'), target = getAttr(el.attrs, 'Target');
    if (!id || !type || !target || out.has(id)) refuse('invalid or duplicate relationship');
    const mode = getAttr(el.attrs, 'TargetMode');
    out.set(id, { type, target, mode, used: false });
  }
  return out;
}
function kind(type) {
  for (const ns of OFFICE) if (type.startsWith(`${ns}/`)) return type.slice(ns.length + 1);
  return null;
}
function resolve(entries, source, rel, expected) {
  if (!rel || kind(rel.type) !== expected || (rel.mode && rel.mode !== 'Internal')) refuse(`invalid ${expected} relationship`);
  const path = targetPath(source, rel.target);
  const entry = entries.get(path);
  if (!entry) refuse(`missing relationship target ${path}`);
  rel.used = true;
  return entry;
}
function apply(text, patches) {
  patches.sort((a, b) => b.start - a.start);
  let last = text.length;
  for (const p of patches) {
    if (p.end > last) refuse('overlapping object edits');
    text = text.slice(0, p.start) + p.value + text.slice(p.end); last = p.start;
  }
  return text;
}
function patchAttrs(text, el, attrs, patches) {
  if (attrs !== el.attrs) patches.push({ start: el.start, end: el.tagEnd, value: `<${el.name}${attrs}${el.selfClosing ? '/>' : '>'}` });
}
function parseCell(value) {
  const m = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(value);
  if (!m) refuse('unsupported table range');
  let col = 0; for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  const row = Number(m[2]);
  if (col > 16384 || row > 1048576) refuse('table range out of bounds');
  return { col, row };
}
function colName(col) { let s = ''; while (col) { col--; s = String.fromCharCode(65 + col % 26) + s; col = Math.floor(col / 26); } return s; }
function parseRange(value) {
  const m = /^([^:]+):([^:]+)$/.exec(value || '');
  if (!m) refuse('unsupported table range');
  const a = parseCell(m[1]), b = parseCell(m[2]);
  if (a.col > b.col || a.row > b.row) refuse('reversed table range');
  return { a, b };
}
function rangeText(r) { return `${colName(r.a.col)}${r.a.row}:${colName(r.b.col)}${r.b.row}`; }
function adjustRange(range, op) {
  const key = op.axis === 'row' ? 'row' : 'col';
  let a = range.a[key], b = range.b[key];
  const end = op.index + op.count - 1;
  if (op.action === 'insert') {
    if (op.index <= a) { a += op.count; b += op.count; }
    else if (op.index <= b) b += op.count;
  } else if (end < a) { a -= op.count; b -= op.count; }
  else if (op.index <= b) {
    const overlap = Math.min(b, end) - Math.max(a, op.index) + 1;
    if (overlap >= b - a + 1) refuse('structural deletion removes an entire table range');
    a = a < op.index ? a : op.index;
    b -= overlap + (op.index < range.a[key] ? range.a[key] - op.index : 0);
  }
  if (a < 1 || b > (key === 'row' ? 1048576 : 16384) || a > b) refuse('object range out of bounds');
  range.a[key] = a; range.b[key] = b;
  return rangeText(range);
}
function tableRewrite(text, op, host, affected) {
  const doc = root(text, 'table', SHEET), patches = [];
  const range = parseRange(getAttr(doc.el.attrs, 'ref'));
  const header = Number(getAttr(doc.el.attrs, 'headerRowCount') ?? '1');
  const totals = Number(getAttr(doc.el.attrs, 'totalsRowCount') ?? '0');
  if (![0, 1].includes(header) || ![0, 1].includes(totals)) refuse('unsupported table header or totals');
  if (affected) {
    const end = op.index + op.count - 1;
    if (op.axis === 'column' && (op.action === 'insert' ? op.index > range.a.col && op.index <= range.b.col : op.index <= range.b.col && end >= range.a.col)) refuse('editing table columns is unsupported');
    if (op.axis === 'row' && op.action === 'delete' && ((header && op.index <= range.a.row && end >= range.a.row) || (totals && op.index <= range.b.row && end >= range.b.row))) refuse('deleting table header or totals is unsupported');
    const oldRef = getAttr(doc.el.attrs, 'ref'), nextRef = adjustRange(range, op);
    if (oldRef !== nextRef) patchAttrs(text, doc.el, setAttr(doc.el.attrs, 'ref', nextRef), patches);
  }
  const direct = children(text, doc.el);
  const allowed = new Set(['autoFilter', 'tableColumns', 'tableStyleInfo']);
  for (const el of direct) {
    if (nsName(el, doc.map).ns !== doc.ns || !allowed.has(localName(el.name))) refuse('unsupported table component');
    if (localName(el.name) === 'autoFilter') {
      if (children(text, el).length) refuse('table filter criteria are unsupported');
      if (affected) {
        const oldRef = getAttr(el.attrs, 'ref'), nextRef = adjustRange(parseRange(oldRef), op);
        if (oldRef !== nextRef) patchAttrs(text, el, setAttr(el.attrs, 'ref', nextRef), patches);
      }
    }
  }
  const columns = direct.filter(el => localName(el.name) === 'tableColumns');
  if (columns.length !== 1) refuse('invalid table columns');
  const columnScope = nsName(columns[0], doc.map).map;
  const colEls = children(text, columns[0]);
  if (colEls.length !== Number(getAttr(columns[0].attrs, 'count')) || colEls.length !== range.b.col - range.a.col + 1) refuse('invalid table column count');
  for (const col of colEls) {
    const scope = nsName(col, columnScope);
    if (localName(col.name) !== 'tableColumn' || scope.ns !== doc.ns) refuse('unsupported table column');
    for (const item of children(text, col)) {
      const name = localName(item.name);
      if (nsName(item, scope.map).ns !== doc.ns || !['calculatedColumnFormula', 'totalsRowFormula'].includes(name) || item.selfClosing || children(text, item).length) refuse('unsupported table column component');
      const formula = unescapeXml(text.slice(item.innerStart, item.innerEnd));
      const translated = rewriteStructuralFormula(formula, { ...op, formulaSheet: host });
      if (!translated.ok) refuse(`table formula: ${translated.reason}`);
      if (formula !== translated.formula) patches.push({ start: item.innerStart, end: item.innerEnd, value: escapeXml(translated.formula) });
    }
  }
  return apply(text, patches);
}
function checkDrawingDependencies(text, el, parent, rels, entries, source, visited) {
  const scope = nsName(el, parent).map;
  const name = localName(el.name);
  if (name === 'extLst' || name === 'f') refuse('unsupported drawing extension or formula');
  for (const attr of attrRecords(el.attrs)) {
    const colon = attr.name.indexOf(':');
    if (colon < 1 || attr.name.startsWith('xmlns:')) continue;
    const ns = scope.get(attr.name.slice(0, colon));
    if (!ns) refuse('undeclared drawing attribute prefix');
    if (!OFFICE.has(ns)) continue;
    if (name !== 'blip' || attr.name.slice(colon + 1) !== 'embed') refuse('unsupported drawing dependency');
    visited.add(resolve(entries, source, rels.get(unescapeXml(attr.value)), 'image').name);
  }
  for (const child of children(text, el)) checkDrawingDependencies(text, child, scope, rels, entries, source, visited);
}

function drawingRewrite(text, op, affected) {
  const doc = root(text, 'wsDr', DRAWING), patches = [];
  for (const anchor of children(text, doc.el)) {
    const info = nsName(anchor, doc.map), name = localName(anchor.name);
    if (info.ns !== doc.ns || !['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(name)) refuse('unsupported drawing anchor');
    const parts = children(text, anchor);
    const names = parts.map(p => localName(p.name));
    if (names.some(n => !['from', 'to', 'pos', 'ext', 'sp', 'pic', 'clientData'].includes(n)) || parts.some(p => nsName(p, info.map).ns !== doc.ns)) refuse('unsupported drawing object');
    if (names.filter(n => n === 'sp' || n === 'pic').length !== 1 || names.filter(n => n === 'clientData').length !== 1) refuse('invalid drawing object');
    const editAs = getAttr(anchor.attrs, 'editAs');
    if (editAs && !['twoCell', 'oneCell', 'absolute'].includes(editAs)) refuse('unsupported drawing editAs');
    if (affected && editAs === 'oneCell') refuse('one-cell drawing layout requires worksheet dimensions');
    const markerNames = name === 'twoCellAnchor' ? ['from', 'to'] : name === 'oneCellAnchor' ? ['from'] : [];
    const positions = [];
    for (const markerName of markerNames) {
      const markers = parts.filter(p => localName(p.name) === markerName);
      if (markers.length !== 1) refuse('invalid drawing marker');
      const marker = markers[0], markerScope = nsName(marker, info.map).map;
      const fields = children(text, marker);
      if (fields.length !== 4 || ['col', 'colOff', 'row', 'rowOff'].some(n => fields.filter(f => localName(f.name) === n && nsName(f, markerScope).ns === doc.ns).length !== 1)) refuse('invalid drawing marker fields');
      for (const field of fields) if (!/^\d+$/.test(text.slice(field.innerStart, field.innerEnd).trim())) refuse('invalid drawing marker value');
      const field = fields.find(f => localName(f.name) === (op.axis === 'row' ? 'row' : 'col'));
      const offset = fields.find(f => localName(f.name) === (op.axis === 'row' ? 'rowOff' : 'colOff'));
      const value = Number(text.slice(field.innerStart, field.innerEnd).trim()), index = op.index - 1;
      let next = value, nextOffset = Number(text.slice(offset.innerStart, offset.innerEnd).trim());
      if (affected && editAs !== 'absolute') {
        if (op.action === 'insert' && value >= index) next += op.count;
        if (op.action === 'delete' && value >= index) {
          next = value >= index + op.count ? value - op.count : index;
          if (value < index + op.count) nextOffset = 0;
        }
      }
      if (!Number.isSafeInteger(next) || next < 0 || next >= (op.axis === 'row' ? 1048576 : 16384) || !Number.isSafeInteger(nextOffset)) refuse('drawing marker out of bounds');
      positions.push([next, nextOffset]);
      if (next !== value) patches.push({ start: field.innerStart, end: field.innerEnd, value: String(next) });
      if (nextOffset !== Number(text.slice(offset.innerStart, offset.innerEnd).trim())) patches.push({ start: offset.innerStart, end: offset.innerEnd, value: String(nextOffset) });
    }
    if (positions.length === 2 && (positions[0][0] > positions[1][0] || (positions[0][0] === positions[1][0] && positions[0][1] >= positions[1][1]))) refuse('structural edit collapses drawing anchor');
  }
  return apply(text, patches);
}

/** Returns a complete, non-mutating set of replacements, or a refusal. */
export function rewriteStructuralObjects(pkg, sheet, op) {
  try {
    if (!pkg || !Array.isArray(pkg.entries) || !sheet?.entry || !sheet.part || !['row', 'column'].includes(op?.axis) || !['insert', 'delete'].includes(op?.action) || !Number.isInteger(op?.index) || !Number.isInteger(op?.count) || op.index < 1 || op.count < 1 || op.index + op.count - 1 > (op.axis === 'row' ? 1048576 : 16384)) refuse('invalid object rewrite request');
    const entries = new Map();
    for (const entry of pkg.entries) { if (entries.has(entry.name)) refuse('duplicate package entry'); entries.set(entry.name, entry); }
    const sheetXml = read(sheet.entry).text, doc = root(sheetXml, 'worksheet', SHEET);
    const rels = relationships(entries, sheet.part), replacements = new Map();
    const visited = new Set([sheet.part]);
    if (entries.has(relsPath(sheet.part))) visited.add(relsPath(sheet.part));
    const affected = sheet.name.toLowerCase() === String(op.sheet).toLowerCase();
    const process = (entry, transform) => {
      const original = read(entry), next = transform(original.text);
      if (next !== original.text) replacements.set(entry.name, { text: next, bom: original.bom });
    };
    for (const child of children(sheetXml, doc.el)) {
      const name = localName(child.name), scope = nsName(child, doc.map);
      if (name !== 'tableParts' && name !== 'drawing') continue;
      if (scope.ns !== doc.ns) refuse('unsupported worksheet object namespace');
      const refs = name === 'drawing' ? [child] : children(sheetXml, child);
      if (name === 'tableParts' && Number(getAttr(child.attrs, 'count')) !== refs.length) refuse('invalid worksheet table count');
      for (const ref of refs) {
        if (name === 'tableParts' && (localName(ref.name) !== 'tablePart' || nsName(ref, scope.map).ns !== doc.ns)) refuse('unsupported worksheet table reference');
        const id = relationshipId(ref, name === 'drawing' ? doc.map : scope.map);
        const entry = resolve(entries, sheet.part, rels.get(id), name === 'drawing' ? 'drawing' : 'table');
        visited.add(entry.name);
        if (name === 'drawing') {
          const drawingRels = relationships(entries, entry.name);
          if (entries.has(relsPath(entry.name))) visited.add(relsPath(entry.name));
          for (const rel of drawingRels.values()) if (kind(rel.type) !== 'image') refuse('unsupported drawing relationship');
          process(entry, text => {
            const drawingDoc = root(text, 'wsDr', DRAWING);
            checkDrawingDependencies(text, drawingDoc.el, new Map(), drawingRels, entries, entry.name, visited);
            for (const rel of drawingRels.values()) if (!rel.used) refuse('unreferenced drawing relationship');
            return drawingRewrite(text, op, affected);
          });
        } else process(entry, text => tableRewrite(text, op, sheet.name, affected));
      }
    }
    for (const rel of rels.values()) {
      if (rel.used) continue;
      const type = kind(rel.type);
      // Hyperlinks carry no cell references in their target; their ranges belong to the sheet.
      if (type === 'hyperlink') continue;
      refuse(`unsupported worksheet relationship: ${type || rel.type}`);
    }
    return { status: 'ok', replacements, visited };
  } catch (error) {
    if (error instanceof Refusal) return { status: 'refused', reason: error.message };
    throw error;
  }
}
