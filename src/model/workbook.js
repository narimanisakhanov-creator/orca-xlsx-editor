// ExcelJS workbook to the view model the grid renders (contract §5). Pure: no DOM, no Orca API,
// so the same code runs in the editor frame and under `node --test`.
//
// The model is sparse on purpose: only cells that carry a value or a style are materialised, and
// rows, columns and merges are read from the structures ExcelJS already holds. Walking the
// rectangle a sheet claims would turn every large, mostly-empty sheet into a stall.

import ExcelJS from 'exceljs'
import { unzipSync } from 'fflate'

import { formatValue, isDateFormat } from './numfmt.js'
import { resolveStyle, themePalette } from './style.js'
import { readZip, writeZip } from '../write/zip.js'
import { findElements, unescapeXml } from '../write/xml.js'

// Excel column width is in characters of the default font; the grid needs pixels.
const COL_WIDTH_TO_PX = 7
const COL_WIDTH_PAD_PX = 5
const MAX_COL_WIDTH_CHARS = 255
const POINTS_TO_PX = 4 / 3

/**
 * @typedef {{ kind: 'number'|'text'|'boolean'|'date'|'error'|'blank',
 *   value: unknown, display: string, formula: string | null, rich?: boolean }} CellValue
 *
 * `rich` marks text that arrived as rich text: the runs are flattened into `value`, the flag keeps
 * the information that the cell had more than one font.
 *
 * @typedef {{ ref: string, row: number, col: number, value: CellValue,
 *   css: { font: object, fill: string | null, border: object, align: object } }} CellView
 *
 * @typedef {{ name: string, index: number, rowCount: number, colCount: number,
 *   cells: Map<string, CellView>, colWidths: Map<number, number>, rowHeights: Map<number, number>,
 *   merges: Array<{ r1: number, c1: number, r2: number, c2: number }>,
 *   freeze: { rows: number, cols: number },
 *   hidden: { rows: Set<number>, cols: Set<number> } }} SheetModel
 *
 * @typedef {{ sheets: SheetModel[], active: number, date1904: boolean, notes: string[] }} WorkbookModel
 */

/**
 * Read a workbook from its bytes.
 *
 * `notes` is additive to the contract: it carries the formatting that could not be represented
 * (unresolvable theme colours, pattern fills drawn solid), de-duplicated across sheets.
 */
export async function loadWorkbook(bytes) {
  const input = bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : bytes
  if (!(input instanceof Uint8Array)) throw new TypeError('loadWorkbook expects a Uint8Array')

  const prepared = normalizeReaderCopy(input)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(prepared.bytes)
  if (prepared.sheetCount != null && workbook.worksheets.length !== prepared.sheetCount) {
    throw new Error(`workbook could not load all sheets (${workbook.worksheets.length}/${prepared.sheetCount})`)
  }
  const palette = themePalette(workbook._themes)
  const date1904 = date1904Of(workbook)
  const notes = new Set(prepared.notes)
  const sheets = workbook.worksheets.map((worksheet, index) => sheetModel(worksheet, index, palette, notes, date1904))
  return {
    sheets,
    active: activeSheet(workbook, sheets.length),
    date1904,
    notes: [...notes]
  }
}

/** Build a reader-only archive copy for ExcelJS's known XML and relationship limitations. */
function normalizeReaderCopy(bytes) {
  const archive = readZip(bytes)
  if (archive.status !== 'ok') throw new Error(`workbook archive cannot be read safely (${archive.status})`)
  const decoder = new TextDecoder()
  const encoder = new TextEncoder()
  const notes = []
  let changed = false
  let sheetCount = null

  for (let index = archive.entries.length - 1; index >= 0; index--) {
    const entry = archive.entries[index]
    if ((entry.name.startsWith('xl/drawings/') && !entry.name.endsWith('.vml')) || entry.name.startsWith('xl/tables/')) {
      archive.entries.splice(index, 1)
      changed = true
    }
  }
  for (const entry of archive.entries) {
    const isWorkbook = entry.name === 'xl/workbook.xml'
    const isWorksheet = entry.name.startsWith('xl/worksheets/') && entry.name.endsWith('.xml')
    const isSharedStrings = entry.name === 'xl/sharedStrings.xml'
    if (!isWorkbook && !isWorksheet && !isSharedStrings) continue
    const text = decoder.decode(entry.data)
    const normalized = normalizeXmlNamespaces(text, { isWorkbook, isWorksheet })
    if (isWorkbook) sheetCount = normalized.sheetCount
    if (normalized.drawings) notes.push('Drawings were not rendered; the original drawing parts will be preserved on save.')
    if (normalized.tables) notes.push('Tables were not rendered; the original table parts will be preserved on save.')
    if (normalized.conditionalFormatting) notes.push('Conditional formatting rules are preserved, but are not displayed.')
    if (normalized.text !== text) {
      entry.data = encoder.encode(normalized.text)
      entry.changed = true
      changed = true
    }
  }

  const relationships = normalizeRelationshipEntries(archive.entries, decoder, encoder)
  changed ||= relationships
  if (sheetCount == null) throw new Error('workbook XML is missing')
  if (!changed) return { bytes, notes, sheetCount }
  try {
    return { bytes: writeZip(archive.entries), notes, sheetCount }
  } catch (error) {
    throw new Error('workbook reader copy could not be prepared', { cause: error })
  }
}

const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships'
])
const SPREADSHEET_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://purl.oclc.org/ooxml/spreadsheetml/main'
])

// Attribute offsets are relative to the token; quoted values may contain `>` and `:`.
function xmlAttributes(token, start) {
  const attributes = []
  const pattern = /([^\s=/>]+)\s*=\s*(["'])(.*?)\2/gs
  pattern.lastIndex = start
  for (let match; (match = pattern.exec(token));) {
    const valueStart = match.index + match[0].indexOf(match[2]) + 1
    attributes.push({ name: match[1], value: match[3], start: match.index, end: match.index + match[1].length, valueStart, valueEnd: valueStart + match[3].length })
  }
  return attributes
}

// ExcelJS ignores CDATA events and expects unprefixed spreadsheet elements and literal `r:id`.
// Rewrite only reader-copy XML tokens, with namespace scope and quote-aware boundaries.
function normalizeXmlNamespaces(text, { isWorkbook, isWorksheet }) {
  const namespaces = [Object.create(null)]
  const elements = []
  const replacements = []
  let sheetCount = 0
  let drawings = false
  let tables = false
  let conditionalFormatting = false
  let offset = 0
  while ((offset = text.indexOf('<', offset)) !== -1) {
    if (text.startsWith('<!--', offset) || text.startsWith('<![CDATA[', offset) || text.startsWith('<?', offset)) {
      const cdata = text.startsWith('<![CDATA[', offset)
      const terminator = cdata ? ']]>' : text.startsWith('<!--', offset) ? '-->' : '?>'
      const end = text.indexOf(terminator, offset + 2)
      if (end < 0) break
      if (cdata) {
        const value = text.slice(offset + 9, end).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        replacements.push({ start: offset, end: end + 3, value })
      }
      offset = end + terminator.length
      continue
    }
    let end = offset + 1
    let quote = null
    for (; end < text.length; end++) {
      const char = text[end]
      if (quote) { if (char === quote) quote = null }
      else if (char === '"' || char === "'") quote = char
      else if (char === '>') break
    }
    if (end >= text.length) break
    const token = text.slice(offset, end + 1)
    const match = /^<(\/?)([^\s/>]+)/.exec(token)
    if (!match || match[2].startsWith('!')) { offset = end + 1; continue }
    const closing = match[1] === '/'
    const name = match[2]
    if (closing) {
      const frame = elements.pop()
      namespaces.pop()
      if (frame?.omit) replacements.push({ start: frame.start, end: end + 1, value: '' })
      if (frame && frame.normalized !== name) replacements.push({ start: offset + 2, end: offset + 2 + name.length, value: frame.normalized })
    } else {
      const scope = Object.create(namespaces[namespaces.length - 1])
      const attrs = xmlAttributes(token, match[0].length)
      for (const attr of attrs) {
        if (attr.name === 'xmlns') scope[''] = unescapeXml(attr.value)
        else if (attr.name.startsWith('xmlns:')) scope[attr.name.slice(6)] = unescapeXml(attr.value)
      }
      const colon = name.indexOf(':')
      const local = colon < 0 ? name : name.slice(colon + 1)
      const uri = scope[colon < 0 ? '' : name.slice(0, colon)]
      const spreadsheet = SPREADSHEET_NAMESPACES.has(uri)
      const normalized = colon >= 0 && spreadsheet ? local : name
      const parent = elements[elements.length - 1]
      const root = elements[0]
      if (isWorkbook && spreadsheet && local === 'sheet' && elements.length === 2 &&
          parent.local === 'sheets' && parent.spreadsheet && root.local === 'workbook' && root.spreadsheet) sheetCount++
      const worksheetChild = isWorksheet && elements.length === 1 && root.local === 'worksheet' && root.spreadsheet && spreadsheet
      const omit = worksheetChild && (local === 'drawing' || local === 'tableParts')
      if (omit && local === 'drawing') drawings = true
      if (omit && local === 'tableParts') tables = true
      if (worksheetChild && local === 'conditionalFormatting') conditionalFormatting = true
      if (isWorksheet && root?.local === 'worksheet' && root.spreadsheet &&
          uri === 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' &&
          local === 'conditionalFormattings') conditionalFormatting = true
      if (normalized !== name) replacements.push({ start: offset + 1, end: offset + 1 + name.length, value: normalized })
      const relationshipId = attrs.find((attr) => {
        const split = attr.name.indexOf(':')
        return split > 0 && attr.name.slice(split + 1) === 'id' && RELATIONSHIP_NAMESPACES.has(scope[attr.name.slice(0, split)])
      })
      const collision = attrs.find((attr) => attr.name === 'r:id' && attr !== relationshipId)
      if (collision) replacements.push({ start: offset + collision.start, end: offset + collision.end, value: '_orca_unrelated_id' })
      if (relationshipId && relationshipId.name !== 'r:id') {
        replacements.push({ start: offset + relationshipId.start, end: offset + relationshipId.end, value: 'r:id' })
      }
      if (!/\/\s*>$/.test(token)) {
        namespaces.push(scope)
        elements.push({ normalized, local, spreadsheet, omit, start: offset })
      } else if (omit) {
        replacements.push({ start: offset, end: end + 1, value: '' })
      }
    }
    offset = end + 1
  }
  replacements.sort((a, b) => a.start - b.start || b.end - a.end)
  const chunks = []
  let cursor = 0
  for (const replacement of replacements) {
    if (replacement.start < cursor) continue
    chunks.push(text.slice(cursor, replacement.start), replacement.value)
    cursor = replacement.end
  }
  chunks.push(text.slice(cursor))
  return { text: chunks.join(''), sheetCount, drawings, tables, conditionalFormatting }
}

function normalizeRelationshipEntries(entries, decoder, encoder) {
  let changed = false
  for (const entry of entries) {
    if (!entry.name.startsWith('xl/worksheets/_rels/') || !entry.name.endsWith('.rels')) continue
    const text = decoder.decode(entry.data)
    const next = rewriteTargets(entry.name, text)
    if (next === text) continue
    entry.data = encoder.encode(next)
    entry.changed = true
    changed = true
  }
  return changed
}

/**
 * The archive rewritten so every worksheet relationship target is relative to the sheet, or `null`
 * when nothing needed rewriting.
 *
 * OPC resolves `Target` against the source part, and a target starting with `/` is an absolute part
 * name. ExcelJS looks worksheet-relationship targets up verbatim for `tableParts`, `comments` and
 * `vmlDrawing`, so `/xl/tables/table1.xml` misses `../tables/table1.xml` and the sheet's
 * `tableParts` entry resolves to `undefined`. Workbook-level targets are left alone: ExcelJS strips
 * a leading `/xl/` from those itself, and rewriting them would rebuild every openpyxl file for
 * nothing. Only the `.rels` parts are inflated here, and the archive is rebuilt only when one of
 * them changed.
 */
export function normalizeAbsoluteTargets(bytes) {
  if (!(bytes instanceof Uint8Array)) return null

  let relationships
  try {
    relationships = unzipSync(bytes, {
      filter: (file) => file.name.endsWith('.rels') && file.name.startsWith('xl/worksheets/_rels/')
    })
  } catch {
    // Not a zip this reader understands: let ExcelJS report it exactly as it would have.
    return null
  }

  const decoder = new TextDecoder()
  const replacements = new Map()
  for (const [name, data] of Object.entries(relationships)) {
    const text = decoder.decode(data)
    // Cheap gate before any parsing: a relationship file without an absolute target has nothing
    // to rewrite, and most workbooks have none.
    if (!text.includes('="/') && !text.includes("='/")) continue
    const next = rewriteTargets(name, text)
    if (next !== text) replacements.set(name, new TextEncoder().encode(next))
  }
  if (replacements.size === 0) return null

  const archive = readZip(bytes)
  if (archive.status !== 'ok') return null
  for (const entry of archive.entries) {
    const replacement = replacements.get(entry.name)
    if (!replacement) continue
    entry.data = replacement
    entry.changed = true
  }

  try {
    return writeZip(archive.entries)
  } catch {
    return null
  }
}

function rewriteTargets(partName, text) {
  const changes = []
  for (const element of findElements(text, 'Relationship')) {
    const token = text.slice(element.start, element.tagEnd)
    const name = /^<([^\s/>]+)/.exec(token)
    if (!name) continue
    const attrs = xmlAttributes(token, name[0].length)
    if (attrs.some((attr) => attr.name === 'TargetMode' && attr.value.toLowerCase() === 'external')) continue
    const target = attrs.find((attr) => attr.name === 'Target')
    if (!target || (!target.value.trim().startsWith('/') && !/[.\\]/.test(target.value))) continue
    changes.push({ start: element.start + target.valueStart, end: element.start + target.valueEnd, value: relativeTarget(partName, target.value) })
  }
  for (let index = changes.length - 1; index >= 0; index--) {
    const change = changes[index]
    text = text.slice(0, change.start) + change.value + text.slice(change.end)
  }
  return text
}

/** `../tables/table1.xml` for `/xl/tables/table1.xml` seen from `xl/worksheets/_rels/sheet1.xml.rels`. */
function relativeTarget(partName, target) {
  const marker = partName.lastIndexOf('_rels/')
  const from = (marker === -1 ? '' : partName.slice(0, marker)).split('/').filter(Boolean)
  const raw = target.trim().replace(/\\/g, '/')
  const targetParts = raw.split('/').filter(Boolean)
  const resolved = raw.startsWith('/') ? [] : from.slice()
  for (const part of targetParts) {
    if (part === '.') continue
    if (part === '..') { if (resolved.length) resolved.pop(); continue }
    resolved.push(part)
  }
  const to = resolved

  let common = 0
  while (common < from.length && common < to.length && from[common] === to[common]) common++
  const steps = []
  for (let i = common; i < from.length; i++) steps.push('..')
  return steps.concat(to.slice(common)).join('/') || '.'
}

/**
 * `true` when the workbook's serial numbers count from 1904-01-01 instead of 1899-12-30. The
 * caller needs it to turn a `Date` back into a serial with the same offset ExcelJS used to read
 * it; the writer itself only ever sees the number.
 *
 * ExcelJS 4.4.0 parses `<workbookPr date1904>` straight into `properties.date1904` as a boolean
 * (`xlsx/xform/book/workbook-properties-xform.js`), so no trip through `workbook.model` is needed
 * — that getter re-serialises every sheet just to reach the same object.
 */
function date1904Of(workbook) {
  return workbook.properties?.date1904 === true
}

function activeSheet(workbook, count) {
  const tab = workbook.views?.[0]?.activeTab
  if (typeof tab === 'number' && tab >= 0 && tab < count) return tab
  return 0
}

function sheetModel(worksheet, index, palette, notes, date1904) {
  const cells = new Map()
  const rowHeights = new Map()
  const hiddenRows = new Set()

  for (const row of rowsOf(worksheet)) {
    const rowIndex = row.number - 1
    if (row.hidden === true) hiddenRows.add(rowIndex)

    // A height of 0 is Excel's way of collapsing a row; the hidden flag is what the grid needs.
    if (typeof row.height === 'number' && row.height > 0) {
      rowHeights.set(rowIndex, Math.round(row.height * POINTS_TO_PX))
    }

    // A sparse array: forEach skips the holes, so only real cells are touched.
    row._cells.forEach((cell) => {
      if (!cell) return
      const view = cellView(cell, palette, notes, date1904)
      if (view) cells.set(`${rowIndex},${cell.col - 1}`, view)
    })
  }

  const { colWidths, hiddenCols } = columnLayout(worksheet)

  return {
    name: worksheet.name,
    index,
    // rowCount/columnCount come from ExcelJS rather than `dimensions`, which reports the extent the
    // writer stored, not the extent that holds data.
    rowCount: worksheet.rowCount,
    colCount: worksheet.columnCount,
    cells,
    colWidths,
    rowHeights,
    merges: mergeRanges(worksheet),
    freeze: freezePane(worksheet),
    hidden: { rows: hiddenRows, cols: hiddenCols }
  }
}

/**
 * The worksheet's rows, sparse.
 *
 * `eachRow()` skips rows that only carry a height or a hidden flag, and `eachRow({includeEmpty})`
 * builds a Row for every index up to the last one, so a sheet whose data ends at row 100,000 would
 * pay for 100,000 objects. The worksheet's own backing array holds exactly the rows that exist.
 */
function rowsOf(worksheet) {
  const rows = worksheet._rows
  if (!Array.isArray(rows)) {
    throw new TypeError('worksheet._rows is not an array: ExcelJS internals moved, review workbook.js')
  }
  const present = []
  for (const row of rows) if (row) present.push(row)
  return present
}

/* ------------------------------------------------------------------ columns --------------------- */

/**
 * Column widths and hidden flags, keyed by 0-based index. `worksheet.columns` is the public list of
 * column objects; a column that was never defined has no width, so it is left out of the maps and
 * the grid falls back to its default.
 */
function columnLayout(worksheet) {
  const colWidths = new Map()
  const hiddenCols = new Set()

  for (const column of worksheet.columns ?? []) {
    const index = column.number - 1
    if (column.hidden === true) hiddenCols.add(index)
    if (typeof column.width === 'number') {
      const chars = clamp(column.width, 0, MAX_COL_WIDTH_CHARS)
      colWidths.set(index, Math.round(chars * COL_WIDTH_TO_PX + COL_WIDTH_PAD_PX))
    }
  }
  return { colWidths, hiddenCols }
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max)
}

/* ------------------------------------------------------------------ merges ---------------------- */

/** Merge references (`A9:E9`) as stored on the sheet; `model.merges` is the same list. */
function mergeRanges(worksheet) {
  const stored = worksheet._merges
  const references = stored && typeof stored === 'object' && !Array.isArray(stored)
    ? Object.values(stored).map((merge) => merge?.range)
    : worksheet.model?.merges

  const ranges = []
  for (const reference of references ?? []) {
    const parsed = parseRange(reference)
    if (parsed) ranges.push(parsed)
  }
  return ranges
}

/** `A9:E9` (or a single `A9`) to 0-based corners. */
function parseRange(text) {
  if (typeof text !== 'string') return null
  const parts = text.split(':')
  const first = parseRef(parts[0])
  const last = parseRef(parts[1] ?? parts[0])
  if (!first || !last) return null
  return {
    r1: Math.min(first.r, last.r),
    c1: Math.min(first.c, last.c),
    r2: Math.max(first.r, last.r),
    c2: Math.max(first.c, last.c)
  }
}

function parseRef(text) {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(text.trim())
  if (!match) return null
  let col = 0
  for (const char of match[1].toUpperCase()) col = col * 26 + (char.charCodeAt(0) - 64)
  return { r: Number(match[2]) - 1, c: col - 1 }
}

/* ------------------------------------------------------------------ view ----------------------- */

function freezePane(worksheet) {
  const view = worksheet.views?.[0]
  if (!view || view.state !== 'frozen') return { rows: 0, cols: 0 }
  return {
    rows: typeof view.ySplit === 'number' ? Math.max(0, Math.trunc(view.ySplit)) : 0,
    cols: typeof view.xSplit === 'number' ? Math.max(0, Math.trunc(view.xSplit)) : 0
  }
}

/* ------------------------------------------------------------------ cells ---------------------- */

/** `CellView` for a cell that is worth drawing, `null` for an empty unstyled cell. */
function cellView(cell, palette, notes, date1904) {
  const value = cellValue(cell, date1904)
  const style = cell.style
  const styled = style !== undefined && style !== null && Object.keys(style).length > 0

  if (value.kind === 'blank' && !styled) return null

  const css = resolveStyle(style, palette)
  for (const note of css.notes) notes.add(note)

  return {
    ref: cell.address,
    row: cell.row - 1,
    col: cell.col - 1,
    value,
    css: { font: css.font, fill: css.fill, border: css.border, align: css.align }
  }
}

/**
 * The cell's value, its Excel-formatted text and the formula behind it.
 *
 * ExcelJS has already turned date-formatted numbers into `Date` objects, so `kind` follows the
 * value's runtime type and only falls back to the format for formula results, which ExcelJS leaves
 * as numbers.
 */
function cellValue(cell, date1904) {
  const numFmt = cell.style?.numFmt

  switch (cell.type) {
    case ExcelJS.ValueType.Null:
    case ExcelJS.ValueType.Merge:
      return blank()

    case ExcelJS.ValueType.Formula: {
      const formula = formulaText(cell)
      const result = cell.result
      if (result === undefined || result === null) {
        // No cached result: a formula-only cell still has to show something, and the formula is the
        // only truth the file carries. Excel would recalculate; nothing here evaluates formulas.
        return { kind: 'text', value: formula, display: formula ?? '', formula }
      }
      const resolved = valueOf(result, numFmt, date1904)
      return { ...resolved, formula }
    }

    case ExcelJS.ValueType.Date:
      return { kind: 'date', value: cell.value, display: formatValue(cell.value, numFmt, true, date1904), formula: null }

    case ExcelJS.ValueType.Number:
      return valueOf(cell.value, numFmt, date1904)

    case ExcelJS.ValueType.Boolean:
      return { kind: 'boolean', value: cell.value, display: cell.value ? 'TRUE' : 'FALSE', formula: null }

    case ExcelJS.ValueType.Error:
      return { kind: 'error', value: cell.value?.error ?? '#VALUE!', display: cell.value?.error ?? '#VALUE!', formula: null }

    case ExcelJS.ValueType.Hyperlink:
      // The hyperlink target is not part of the view model; the visible text is.
      return textValue(String(cell.value?.text ?? cell.text ?? ''), false)

    case ExcelJS.ValueType.RichText:
      return textValue(richText(cell.value), true)

    case ExcelJS.ValueType.SharedString:
    case ExcelJS.ValueType.String:
      return textValue(typeof cell.value === 'string' ? cell.value : String(cell.value ?? ''), false)

    default:
      return valueOf(cell.value, numFmt, date1904)
  }
}

function blank() {
  return { kind: 'blank', value: null, display: '', formula: null }
}

function valueOf(value, numFmt, date1904) {
  if (value instanceof Date) {
    return { kind: 'date', value, display: formatValue(value, numFmt, true, date1904), formula: null }
  }
  if (typeof value === 'number') {
    const date = isDateFormat(numFmt)
    return {
      kind: date ? 'date' : 'number',
      value,
      display: formatValue(value, numFmt, date, date1904),
      formula: null
    }
  }
  if (typeof value === 'boolean') {
    return { kind: 'boolean', value, display: value ? 'TRUE' : 'FALSE', formula: null }
  }
  if (value && typeof value === 'object' && typeof value.error === 'string') {
    return { kind: 'error', value: value.error, display: value.error, formula: null }
  }
  if (value === null || value === undefined) return blank()
  return textValue(String(value), false)
}

function textValue(text, rich) {
  const cell = { kind: 'text', value: text, display: text, formula: null }
  if (rich) cell.rich = true
  return cell
}

/** Rich text keeps its characters and loses the runs; `CellValue.rich` records that it had them. */
function richText(value) {
  const runs = value?.richText
  if (!Array.isArray(runs)) return String(value ?? '')
  return runs.map((run) => run.text ?? '').join('')
}

function formulaText(cell) {
  // `cell.formula` is translated for shared-formula slaves; the raw model may only carry the master.
  const formula = typeof cell.formula === 'string' && cell.formula !== '' ? cell.formula : cell.model?.formula
  if (!formula) return null
  return formula.startsWith('=') ? formula.slice(1) : formula
}
