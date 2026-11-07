// Atomic structural edits over original OPC entries. Unsupported constructs are refused.
import { xlsxPackage } from './xlsx-writer.js'
import { findElement, findElements, getAttr, removeAttr, replaceElements, setAttr, topLevelElements, unescapeXml, escapeXml } from './xml.js'
import { rewriteStructuralFormula } from './structural-formula.js'
import { rewriteStructuralObjects } from './structural-objects.js'

const MAX_ROW = 1048576
const MAX_COL = 16384
const encoder = new TextEncoder()
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet'
class Refusal extends Error {}
const fail = (message) => { throw new Refusal(message) }

function decode(entry) {
  const bytes = entry?.data
  if (!(bytes instanceof Uint8Array)) fail(`part ${entry?.name} is unreadable`)
  if (bytes[0] === 0xff && bytes[1] === 0xfe || bytes[0] === 0xfe && bytes[1] === 0xff) fail(`part ${entry.name} is UTF-16`)
  const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
  try { return { text: new TextDecoder('utf-8', { fatal: true }).decode(bom ? bytes.subarray(3) : bytes), bom } }
  catch { fail(`part ${entry.name} is not UTF-8`) }
}
function encode(text, bom) {
  const bytes = encoder.encode(text)
  if (!bom) return bytes
  const out = new Uint8Array(bytes.length + 3)
  out.set([0xef, 0xbb, 0xbf]); out.set(bytes, 3)
  return out
}
function children(xml, element) { return element.selfClosing ? [] : topLevelElements(xml.slice(element.start, element.end)) }
function checkChildren(xml, allowed, label) {
  for (const child of topLevelElements(xml)) if (!allowed.has(child.local)) fail(`${label} contains unsupported ${child.local}`)
}
function retag(xml, el, attrs, inner = null) {
  if (el.selfClosing) return `<${el.name}${attrs}/>`
  return `<${el.name}${attrs}>${inner === null ? xml.slice(el.innerStart, el.innerEnd) : inner}</${el.name}>`
}
function colNumber(s) {
  let n = 0
  for (const c of s.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64
  return n
}
function colName(n) {
  let s = ''
  while (n) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26) }
  return s
}
function ref(text) {
  const m = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(text ?? '')
  if (!m) fail(`unsupported cell reference ${text}`)
  const c = colNumber(m[1]), r = Number(m[2])
  if (c > MAX_COL || r > MAX_ROW) fail(`cell reference outside sheet ${text}`)
  return { r, c }
}
function refText(p) { return `${colName(p.c)}${p.r}` }
function move(n, op, max) {
  if (op.action === 'insert') {
    const result = n >= op.index ? n + op.count : n
    if (result > max) fail('structural insertion exceeds worksheet bounds')
    return result
  }
  if (n < op.index) return n
  if (n < op.index + op.count) return null
  return n - op.count
}
function moveRef(value, op) {
  const p = ref(value), key = op.axis === 'row' ? 'r' : 'c'
  p[key] = move(p[key], op, key === 'r' ? MAX_ROW : MAX_COL)
  return p[key] === null ? null : refText(p)
}
function moveRange(value, op) {
  const parts = String(value ?? '').split(':')
  if (parts.length > 2) fail(`unsupported range ${value}`)
  const a = ref(parts[0]), b = ref(parts[parts.length - 1])
  if (a.r > b.r || a.c > b.c) fail(`invalid range ${value}`)
  const key = op.axis === 'row' ? 'r' : 'c', max = key === 'r' ? MAX_ROW : MAX_COL
  let lo = a[key], hi = b[key]
  if (op.action === 'insert') {
    if (lo >= op.index) lo += op.count
    if (hi >= op.index) hi += op.count
    if (hi > max) fail('range exceeds worksheet bounds')
  } else {
    const end = op.index + op.count - 1
    if (lo > end) { lo -= op.count; hi -= op.count }
    else if (hi >= op.index) {
      if (lo >= op.index && hi <= end) return null
      hi = hi > end ? hi - op.count : op.index - 1
      if (lo >= op.index) lo = op.index
    }
  }
  a[key] = lo; b[key] = hi
  return parts.length === 1 ? refText(a) : `${refText(a)}:${refText(b)}`
}
/** Active cell after the edit: a deleted cell moves to the cell that now occupies it. */
function movePoint(value, op) {
  const single = String(value ?? '').trim()
  const updated = moveRange(single, op)
  if (updated !== null) return updated.split(':')[0]
  const match = /^([A-Za-z]{1,3})([1-9][0-9]*)$/.exec(single)
  if (!match) fail(`unsupported reference ${value}`)
  const c = colNumber(match[1]), r = Number(match[2])
  const max = op.axis === 'row' ? MAX_ROW : MAX_COL
  const index = Math.min(op.index, max)
  if (op.axis === 'row') return refText({ r: r < op.index ? r : index, c })
  return refText({ r, c: c < op.index ? c : index })
}
/** A space-separated Excel reference list (`sqref`); deleted entries disappear. */
function moveSqref(value, op) {
  const parts = String(value ?? '').trim().split(/\s+/).filter(Boolean)
  const out = []
  for (const part of parts) {
    const updated = moveRange(part, op)
    if (updated !== null) out.push(updated)
  }
  return out.join(' ')
}
/** Frozen/split pane boundary: rows or columns above it. */
function moveSplit(value, op) {
  const n = Number(value)
  if (!Number.isFinite(n) || n < 0) fail('invalid pane split')
  if (op.action === 'insert') return n >= op.index ? n + op.count : n
  const end = op.index + op.count - 1
  if (n < op.index) return n
  return n <= end ? op.index - 1 : n - op.count
}
function transformViews(xml, op) {
  let next = replaceElements(xml, 'pane', (el) => {
    let attrs = el.attrs
    const key = op.axis === 'row' ? 'ySplit' : 'xSplit'
    if (getAttr(attrs, key) !== null) attrs = setAttr(attrs, key, String(moveSplit(getAttr(attrs, key), op)))
    const topLeft = getAttr(attrs, 'topLeftCell')
    if (topLeft !== null) attrs = setAttr(attrs, 'topLeftCell', movePoint(topLeft, op))
    return retag(xml, el, attrs)
  })
  next = replaceElements(next, 'selection', (el) => {
    let attrs = el.attrs
    const active = getAttr(attrs, 'activeCell')
    if (active !== null) attrs = setAttr(attrs, 'activeCell', movePoint(active, op))
    const sqref = getAttr(attrs, 'sqref')
    if (sqref !== null) {
      const moved = moveSqref(sqref, op)
      if (!moved) return ''
      attrs = setAttr(attrs, 'sqref', moved)
    }
    return retag(next, el, attrs)
  })
  return next
}
/** Column definitions shift, shrink or split when the edited sheet is a column edit. */
function transformColumns(xml, op) {
  if (op.axis !== 'column') return xml
  const container = findElement(xml, 'cols')
  if (!container) return xml
  const whole = xml.slice(container.start, container.end)
  const columns = findElements(whole, 'col')
  const pieces = []
  for (const col of columns) {
    const min = Number(getAttr(col.attrs, 'min'))
    const max = Number(getAttr(col.attrs, 'max') ?? getAttr(col.attrs, 'min'))
    if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max > MAX_COL || min > max) fail('invalid column definition')
    const push = (a, b, attrs) => pieces.push({ a, b, attrs: setAttr(setAttr(attrs, 'min', a), 'max', b) })
    if (op.action === 'insert') {
      const a = min >= op.index ? min + op.count : min
      const b = max >= op.index ? max + op.count : max
      if (b > MAX_COL) fail('column definition exceeds worksheet bounds')
      push(a, b, col.attrs)
      continue
    }
    const end = op.index + op.count - 1
    if (max < op.index) push(min, max, col.attrs)
    else if (min > end) push(min - op.count, max - op.count, col.attrs)
    else {
      if (min < op.index) push(min, op.index - 1, col.attrs)
      if (max > end) push(op.index, max - op.count, col.attrs)
    }
  }
  if (pieces.length === 0) return xml.slice(0, container.start) + xml.slice(container.end)
  const name = columns[0]?.name ?? 'col'
  const inner = pieces.map((piece) => `<${name}${piece.attrs}/>`).join('')
  return xml.slice(0, container.start) + `<${container.name}${container.attrs}>${inner}</${container.name}>` + xml.slice(container.end)
}
function formulaText(xml, el, op, formulaSheet) {
  if (el.selfClosing || getAttr(el.attrs, 't') && getAttr(el.attrs, 't') !== 'normal') fail('shared, array or empty formula is unsupported')
  const source = xml.slice(el.innerStart, el.innerEnd)
  if (source.includes('<')) fail('formula contains unsupported XML markup')
  const result = rewriteStructuralFormula(unescapeXml(source), { ...op, formulaSheet })
  if (!result.ok) fail(result.reason || 'unsupported formula')
  return retag(xml, el, el.attrs, escapeXml(result.formula))
}
function transformSheet(xml, name, op, target) {
  const allowed = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'mergeCells', 'autoFilter', 'hyperlinks', 'conditionalFormatting', 'dataValidations', 'pageMargins', 'pageSetup', 'printOptions', 'headerFooter', 'drawing', 'tableParts'])
  checkChildren(xml, allowed, `sheet ${name}`)
  if (findElements(xml, 'extLst').length || findElements(xml, 'sheetProtection').length) fail(`sheet ${name} has protection or extensions`)
  const data = topLevelElements(xml).filter((el) => el.local === 'sheetData')
  if (data.length !== 1) fail(`sheet ${name} has no unique sheetData`)
  let next = replaceElements(xml, 'f', (el) => formulaText(xml, el, op, name))
  if (!target) return next
  const dataEl = topLevelElements(next).find((el) => el.local === 'sheetData')
  const dataXml = next.slice(dataEl.start, dataEl.end)
  checkChildren(dataXml, new Set(['row']), 'sheetData')
  let previousRow = MAX_ROW + 1
  const rewritten = replaceElements(dataXml, 'row', (row, rowXml) => {
    const rowNumber = Number(getAttr(row.attrs, 'r'))
    if (!Number.isInteger(rowNumber) || rowNumber >= previousRow || rowNumber < 1) fail('row numbers are missing, unordered or invalid')
    previousRow = rowNumber
    checkChildren(rowXml, new Set(['c']), 'row')
    const newRow = op.axis === 'row' ? move(rowNumber, op, MAX_ROW) : rowNumber
    if (newRow === null) return ''
    let previousCol = MAX_COL + 1
    const cells = replaceElements(rowXml, 'c', (cell, cellXml) => {
      const oldRef = getAttr(cell.attrs, 'r'), p = ref(oldRef)
      if (p.r !== rowNumber || p.c >= previousCol) fail('cell coordinates are inconsistent or unordered')
      previousCol = p.c
      if (getAttr(cell.attrs, 'cm') !== null || getAttr(cell.attrs, 'vm') !== null || findElement(cellXml, 'extLst')) fail('cell metadata is unsupported')
      checkChildren(cellXml, new Set(['f', 'v', 'is']), 'cell')
      const newRef = moveRef(oldRef, op)
      if (newRef === null) return ''
      return retag(cellXml, findElement(cellXml, 'c'), setAttr(cell.attrs, 'r', newRef))
    })
    const local = findElement(cells, 'row')
    let attrs = setAttr(local.attrs, 'r', String(newRow))
    attrs = removeAttr(attrs, 'spans')
    return retag(cells, local, attrs)
  })
  next = next.slice(0, dataEl.start) + rewritten + next.slice(dataEl.end)
  next = replaceElements(next, 'mergeCell', (el) => {
    const updated = moveRange(getAttr(el.attrs, 'ref'), op)
    return updated === null ? '' : retag(next, el, setAttr(el.attrs, 'ref', updated))
  })
  // Dimensions are advisory; removing the old bound is safer than preserving a stale one.
  next = replaceElements(next, 'dimension', () => '')
  // Empty merges are not valid. Update the count after removing deleted ranges.
  next = replaceElements(next, 'mergeCells', (el, whole) => {
    const count = findElements(whole, 'mergeCell').length
    return count ? retag(next, el, setAttr(el.attrs, 'count', String(count))) : ''
  })
  next = transformColumns(next, op)
  next = transformViews(next, op)
  next = replaceElements(next, 'autoFilter', (el) => {
    const updated = moveRange(getAttr(el.attrs, 'ref'), op)
    return updated === null ? '' : retag(next, el, setAttr(el.attrs, 'ref', updated))
  })
  next = replaceElements(next, 'hyperlink', (el) => {
    const updated = moveRange(getAttr(el.attrs, 'ref'), op)
    return updated === null ? '' : retag(next, el, setAttr(el.attrs, 'ref', updated))
  })
  next = replaceElements(next, 'conditionalFormatting', (el) => {
    const moved = moveSqref(getAttr(el.attrs, 'sqref'), op)
    return moved ? retag(next, el, setAttr(el.attrs, 'sqref', moved)) : ''
  })
  next = replaceElements(next, 'dataValidation', (el) => {
    const moved = moveSqref(getAttr(el.attrs, 'sqref'), op)
    return moved ? retag(next, el, setAttr(el.attrs, 'sqref', moved)) : ''
  })
  next = replaceElements(next, 'dataValidations', (el, whole) => {
    const count = findElements(whole, 'dataValidation').length
    return count ? retag(next, el, setAttr(el.attrs, 'count', String(count))) : ''
  })
  return next
}
function transformBook(xml, op, sheets) {
  const allowed = new Set(['fileVersion', 'workbookPr', 'bookViews', 'sheets', 'definedNames', 'calcPr'])
  checkChildren(xml, allowed, 'workbook')
  if (findElement(xml, 'workbookProtection') || findElement(xml, 'extLst')) fail('protected or extended workbook')
  let next = replaceElements(xml, 'definedName', (el) => {
    const local = getAttr(el.attrs, 'localSheetId')
    let formulaSheet = null
    if (local !== null) {
      const n = Number(local)
      if (!Number.isInteger(n) || n < 0 || n >= sheets.length) fail('invalid local defined-name scope')
      formulaSheet = sheets[n].name
    }
    return formulaText(xml, el, op, formulaSheet)
  })
  const calc = topLevelElements(next).find((el) => el.local === 'calcPr')
  if (calc) next = next.slice(0, calc.start) + retag(next, calc, setAttr(calc.attrs, 'fullCalcOnLoad', '1')) + next.slice(calc.end)
  else {
    const root = findElement(next, 'workbook')
    const prefix = root.name.includes(':') ? root.name.slice(0, root.name.indexOf(':') + 1) : ''
    next = next.slice(0, root.innerEnd) + `<${prefix}calcPr fullCalcOnLoad="1"/>` + next.slice(root.innerEnd)
  }
  return next
}

const REL_TYPES = new Set([
  'officeDocument', 'worksheet', 'styles', 'sharedStrings', 'theme', 'drawing', 'table', 'image', 'hyperlink',
  'extended-properties', 'custom-properties', 'vbaProject', 'calcChain',
])
function preflightPackage(pkg) {
  const names = new Set()
  for (const entry of pkg.entries) {
    if (!entry || typeof entry.name !== 'string') fail('invalid package entry')
    const key = entry.name.toLowerCase()
    if (names.has(key)) fail(`duplicate package part ${entry.name}`)
    names.add(key)
    if (!entry.name.endsWith('.rels')) continue
    const xml = decode(entry).text
    checkChildren(xml, new Set(['Relationship']), entry.name)
    for (const rel of findElements(xml, 'Relationship')) {
      const type = getAttr(rel.attrs, 'Type') ?? ''
      const match = /^(?:http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships|http:\/\/purl\.oclc\.org\/ooxml\/officeDocument\/relationships)\/([^/]+)$/.exec(type)
      const core = type === 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
      if (!core && (!match || !REL_TYPES.has(match[1]))) fail(`unsupported relationship type ${type}`)
      if ((getAttr(rel.attrs, 'TargetMode') ?? '').toLowerCase() === 'external') fail('external relationships are unsupported')
    }
  }
}
function isCorePart(name, pkg) {
  if (name.endsWith('/')) return true
  if (name === pkg.book.path || name === pkg.book.relsPath || name === '[Content_Types].xml' || name === '_rels/.rels') return true
  return /^(?:docProps\/(?:app|core|custom)\.xml|xl\/(?:styles|sharedStrings)\.xml|xl\/theme\/theme[0-9]+\.xml|xl\/vbaProject\.bin)$/i.test(name)
}

/** Apply a 1-based insert/delete to a workbook, or refuse without changing any input entry. */
export function applyStructural(entries, operation) {
  try {
    if (!operation || typeof operation.sheet !== 'string' || !operation.sheet || !['row', 'column'].includes(operation.axis) || !['insert', 'delete'].includes(operation.action)) fail('invalid structural operation')
    const max = operation.axis === 'row' ? MAX_ROW : MAX_COL
    if (!Number.isInteger(operation.index) || !Number.isInteger(operation.count) || operation.index < 1 || operation.count < 1 || operation.index + operation.count - 1 > max) fail('structural range outside worksheet bounds')
    const diagnostic = {}
    const pkg = xlsxPackage(entries, diagnostic)
    if (!pkg) fail(diagnostic.reason || 'unreadable workbook')
    preflightPackage(pkg)
    if (pkg.book.sheets.filter((s) => s.name.toLowerCase() === operation.sheet.toLowerCase()).length !== 1) fail('target sheet is missing or ambiguous')
    const calcParts = pkg.entries.filter((entry) => /(^|\/)calcChain\.xml$/i.test(entry.name))
    const dropped = new Set(calcParts.map((entry) => entry.name))
    const target = pkg.book.sheets.find((s) => s.name.toLowerCase() === operation.sheet.toLowerCase())
    const op = { ...operation, sheet: target.name }
    const replacements = new Map()
    const visited = new Set()
    const book = transformBook(pkg.book.text, op, pkg.book.sheets)
    if (book !== pkg.book.text) replacements.set(pkg.book.entry.name, { text: book, bom: pkg.book.bom })
    for (const sheet of pkg.book.sheets) {
      if (!sheet.entry || ![REL, STRICT_REL].includes(sheet.relType)) fail(`sheet ${sheet.name} is not a supported worksheet`)
      const source = decode(sheet.entry)
      const changed = transformSheet(source.text, sheet.name, op, sheet === target)
      if (changed !== source.text) replacements.set(sheet.entry.name, { text: changed, bom: source.bom })
      const objects = rewriteStructuralObjects(pkg, sheet, op)
      if (objects.status !== 'ok') fail(objects.reason || 'unsupported related workbook object')
      if (!(objects.visited instanceof Set)) fail('object traversal did not identify inspected parts')
      for (const name of objects.visited) visited.add(name)
      for (const [name, value] of objects.replacements) {
        if (replacements.has(name)) fail(`conflicting structural changes to ${name}`)
        replacements.set(name, value)
      }
    }
    // A stale calc chain after moved cells makes Excel show wrong results; Excel rebuilds it.
    // The relationship and content-type override must go with the part.
    if (calcParts.length) {
      if (pkg.book.relsEntry) {
        const source = decode(pkg.book.relsEntry)
        const next = replaceElements(source.text, 'Relationship', (el) =>
          (getAttr(el.attrs, 'Type') ?? '').endsWith('/calcChain') ? '' : null)
        if (next !== source.text) replacements.set(pkg.book.relsEntry.name, { text: next, bom: source.bom })
      }
      const contentTypes = entries.find((entry) => entry.name === '[Content_Types].xml')
      if (contentTypes) {
        const source = decode(contentTypes)
        const next = replaceElements(source.text, 'Override', (el) =>
          /(^|\/)calcChain\.xml$/i.test(getAttr(el.attrs, 'PartName') ?? '') ? '' : null)
        if (next !== source.text) replacements.set(contentTypes.name, { text: next, bom: source.bom })
      }
    }
    for (const entry of entries) {
      if (dropped.has(entry.name)) continue
      if (!isCorePart(entry.name, pkg) && !visited.has(entry.name)) fail(`uninspected workbook part ${entry.name}`)
    }
    return { status: 'ok', entries: entries.filter((entry) => !dropped.has(entry.name)).map((entry) => {
      const replacement = replacements.get(entry.name)
      return replacement ? { ...entry, data: encode(replacement.text, replacement.bom), changed: true } : entry
    }) }
  } catch (error) {
    if (error instanceof Refusal) return { status: 'refused', reason: error.message }
    throw error
  }
}
