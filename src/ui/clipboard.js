// Clipboard text for the grid: TSV, CSV and Markdown built from a SheetModel.
// The formatters are pure functions (no DOM, no Orca API) so `node --test` runs them directly;
// only copyText touches the DOM, because the editor frame has no navigator.clipboard.

const TAB_OR_NEWLINE = /[\t\r\n]+/g

/**
 * A selection larger than this is refused for HTML rather than turned into an unbounded string.
 * Clipboard text stays uncapped; only the richer format has a hard ceiling, and the error is
 * explicit so a caller can fall back to TSV instead of pasting a silently cut table.
 */
export const MAX_HTML_CELLS = 250000

/**
 * Serialized UTF-8 ceiling. Cell count alone does not bound memory: 250,000 cells can each hold a
 * 32,000-character shared string, and one pasted table would try to build gigabytes. The
 * serializer refuses once the escaped markup, text and inline styles cross this.
 */
export const MAX_HTML_BYTES = 8 * 1024 * 1024

/** Browser ceilings for table spans; a larger merge cannot be represented in one cell. */
export const MAX_HTML_COLSPAN = 1000
export const MAX_HTML_ROWSPAN = 65534

/**
 * Excel treats a cell starting with `=`, `+`, `-`, `@`, a tab or a carriage return as a formula, so
 * exported text that starts that way is prefixed with an apostrophe. Plain numbers (`-2`, `+5`) are
 * left alone: they are values, not formulas, and turning them into text would be worse.
 */
function defuse(text) {
  if (typeof text !== 'string' || text === '') return text
  if (/^[+-]?\d+([.,]\d+)?$/.test(text)) return text
  return /^[=+\-@\t\r]/.test(text) ? `'${text}` : text
}

/** Cells are keyed `row,col`, 0-based, exactly as the model stores them. */
function cellAt(sheet, r, c) {
  return sheet?.cells?.get(`${r},${c}`) ?? null
}

/** What Excel would show, or `=FORMULA` when the caller asked for formulas. */
function cellText(sheet, r, c, formulas) {
  const value = cellAt(sheet, r, c)?.value
  if (!value) return ''
  if (formulas && value.formula) return `=${value.formula}`
  return value.display ?? ''
}

/** Whether the text at this cell is a real formula the caller asked to export as one. */
function isFormulaAt(sheet, r, c, formulas) {
  return Boolean(formulas && cellAt(sheet, r, c)?.value?.formula)
}

/** Inclusive box of the selection, or null when there is nothing to copy. */
function bounds(selection) {
  if (!selection) return null
  const r1 = Number(selection.r1)
  const c1 = Number(selection.c1)
  const r2 = Number(selection.r2)
  const c2 = Number(selection.c2)
  if (![r1, c1, r2, c2].every(Number.isInteger)) return null
  if (r2 < r1 || c2 < c1) return null
  return { r1, c1, r2, c2 }
}

/**
 * Conservative upper bound for a Markdown or fenced TSV handoff, without building a matrix or
 * expanding repeated shared strings. Four units per source unit cover Markdown escaping and two
 * TSV fences even in the worst case; per-cell overhead covers separators, rules and defusing.
 * The caller reserves space for its own context header separately.
 */
export function selectionFitsTextBudget(sheet, selection, { formulas = false, maxChars } = {}) {
  const box = bounds(selection)
  if (!sheet || !box || !Number.isFinite(maxChars) || maxChars < 0) return false
  let total = 0
  for (let r = box.r1; r <= box.r2; r++) {
    for (let c = box.c1; c <= box.c2; c++) {
      const value = cellAt(sheet, r, c)?.value
      const length = formulas && value?.formula ? value.formula.length + 1 : (value?.display ?? '').length
      total += length * 4 + 16
      if (total > maxChars) return false
    }
  }
  return true
}

function textMatrix(sheet, box, formulas) {
  if (!sheet || !box) return []
  const rows = []
  for (let r = box.r1; r <= box.r2; r++) {
    const row = []
    for (let c = box.c1; c <= box.c2; c++) {
      const text = cellText(sheet, r, c, formulas)
      // A formula the caller asked for is exported as a formula; only cell text that merely looks
      // like one is defused.
      row.push({ text, formula: isFormulaAt(sheet, r, c, formulas) })
    }
    rows.push(row)
  }
  return rows
}

/**
 * TSV as Excel puts it on the clipboard: `\n` between rows, tabs between cells. TSV has no
 * quoting, so a tab or line break inside a value cannot survive as itself — it becomes a space.
 * That is deliberate: one cell stays one field and a pasted block keeps its shape instead of
 * the value splitting into extra cells. CSV is the format that escapes.
 */
export function selectionToTsv(sheet, selection, { formulas = false } = {}) {
  const box = bounds(selection)
  return textMatrix(sheet, box, formulas)
    .map((row) => row.map((cell) => (cell.formula ? cell.text : defuse(cell.text.replace(TAB_OR_NEWLINE, ' ')))).join('\t'))
    .join('\n')
}

/** RFC 4180 field: quoted only when it holds a comma, a quote or a line break. */
function csvField(cell) {
  const safe = cell.formula ? cell.text : defuse(cell.text)
  if (!/[",\r\n]/.test(safe)) return safe
  return `"${safe.replace(/"/g, '""')}"`
}

/** RFC 4180: CRLF between records, doubled quotes, quoting only where the grammar requires it. */
export function selectionToCsv(sheet, selection, { formulas = false } = {}) {
  const box = bounds(selection)
  return textMatrix(sheet, box, formulas)
    .map((row) => row.map(csvField).join(','))
    .join('\r\n')
}

/** A column is right-aligned when every non-empty data cell in it is a number. */
function columnIsNumeric(sheet, box, c, formulas, fromRow) {
  let seen = false
  for (let r = fromRow; r <= box.r2; r++) {
    const value = cellAt(sheet, r, c)?.value
    if (!value || value.kind === 'blank') continue
    seen = true
    if (value.kind !== 'number' || (formulas && value.formula)) return false
  }
  return seen
}

/** GFM cell text: pipes are escaped and line breaks become `<br>`, or the table would break. */
function markdownCell(text) {
  return text.replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')
}

/**
 * GFM table. The first row of a range becomes the header; GFM has no headerless table, so a
 * one-row selection is written as its own header row.
 */
export function selectionToMarkdown(sheet, selection, { formulas = false } = {}) {
  const box = bounds(selection)
  if (!box || !sheet) return ''
  const rows = textMatrix(sheet, box, formulas)
  if (!rows.length) return ''
  const header = rows[0]
  const body = rows.length > 1 ? rows.slice(1) : []
  const dataFrom = rows.length > 1 ? box.r1 + 1 : box.r1
  const cellOf = (cell) => (typeof cell === 'string' ? cell : cell.formula ? cell.text : defuse(cell.text))
  const line = (cells) => `| ${cells.map((cell) => markdownCell(cellOf(cell))).join(' | ')} |`
  const rule = rows[0].map((_, index) => {
    const column = box.c1 + index
    return columnIsNumeric(sheet, box, column, formulas, dataFrom) ? '---:' : '---'
  })
  return [line(header), line(rule), ...body.map(line)].join('\n')
}

/* ------------------------------------------------------------------ html ------------------------ */

/** Text in a table cell, escaped, with Excel's line breaks kept as `<br>`. */
function htmlText(text) {
  return escapeHtml(text).replace(/\r\n|\r|\n/g, '<br>')
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/** A workbook can name any font; keep it a font name and out of the rest of the stylesheet. */
function fontFamily(name) {
  return String(name).replace(/['"\\\r\n;]/g, '').slice(0, 120)
}

/** Bytes of `<tr></tr>`, the markup a row adds before its cells. */
const ROW_TAG_BYTES = 9

/** UTF-8 byte length without allocating an encoded copy of the string. */
function utf8Length(text) {
  let bytes = 0
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index)
    if (code < 0x80) bytes += 1
    else if (code < 0x800) bytes += 2
    else if (code >= 0xd800 && code <= 0xdbff) {
      bytes += 4
      index++
    } else bytes += 3
  }
  return bytes
}

/** Inline CSS for one `CellView.css`, or `''` when there is nothing Excel would keep. */
function htmlStyle(css, value) {
  const declarations = []
  const font = css?.font ?? {}
  const align = css?.align ?? {}
  const border = css?.border ?? {}

  if (font.bold) declarations.push('font-weight:700')
  if (font.italic) declarations.push('font-style:italic')
  if (font.size != null && Number.isFinite(Number(font.size))) {
    declarations.push(`font-size:${Math.round(Number(font.size) * 100) / 100}pt`)
  }
  if (font.color) declarations.push(`color:${font.color}`)
  if (font.family && fontFamily(font.family)) declarations.push(`font-family:'${fontFamily(font.family)}'`)

  const decorations = []
  if (font.underline) decorations.push('underline')
  if (font.strike) decorations.push('line-through')
  if (decorations.length) declarations.push(`text-decoration:${decorations.join(' ')}`)
  if (font.underline === 'double') declarations.push('text-decoration-style:double')

  const fill = typeof css?.fill === 'string' ? css.fill : null
  if (fill) declarations.push(`background-color:${fill}`)

  const horizontal = String(align.horizontal ?? '').toLowerCase()
  if (horizontal === 'center' || horizontal === 'centercontinuous' || horizontal === 'distributed') {
    declarations.push('text-align:center')
  } else if (horizontal === 'right' || horizontal === 'fill') {
    declarations.push('text-align:right')
  } else if (horizontal === 'justify') {
    declarations.push('text-align:justify')
  } else if (horizontal === 'left') {
    declarations.push('text-align:left')
  } else if (value?.kind === 'number' || value?.kind === 'date') {
    // The grid right-aligns numbers that carry no explicit alignment; the paste should match it.
    declarations.push('text-align:right')
  }

  const vertical = String(align.vertical ?? '').toLowerCase()
  if (vertical === 'top') declarations.push('vertical-align:top')
  else if (vertical === 'bottom') declarations.push('vertical-align:bottom')
  else if (vertical === 'center' || vertical === 'middle') declarations.push('vertical-align:middle')

  const indent = Number(align.indent)
  if (Number.isFinite(indent) && indent > 0) declarations.push(`padding-left:${6 + indent * 8}px`)
  if (align.wrapText) declarations.push('white-space:normal', 'overflow-wrap:anywhere')
  if (font.vertAlign === 'superscript') declarations.push('vertical-align:super')
  else if (font.vertAlign === 'subscript') declarations.push('vertical-align:sub')

  for (const side of ['top', 'right', 'bottom', 'left']) {
    const edge = border?.[side]
    if (!edge?.style) continue
    const name = String(edge.style).toLowerCase()
    const style = name === 'dashed' || name === 'dotted' || name === 'double' ? name : 'solid'
    const width = name === 'medium' || name === 'thick' ? 2 : 1
    declarations.push(`border-${side}:${width}px ${style} ${edge.color || '#000000'}`)
  }

  return declarations.join(';')
}

/** One `<td>`, with the span and the content pulled from the merge's real anchor. */
function htmlCell(view, span, formulas) {
  const value = view?.value
  let text = ''
  let formula = false
  if (value) {
    formula = Boolean(formulas && value.formula)
    text = formula ? `=${value.formula}` : value.display ?? ''
    if (!formula) text = defuse(text)
  }

  const attributes = []
  if (span?.colspan > 1) attributes.push(`colspan="${span.colspan}"`)
  if (span?.rowspan > 1) attributes.push(`rowspan="${span.rowspan}"`)
  const style = htmlStyle(view?.css, value)
  if (style) attributes.push(`style="${escapeHtml(style)}"`)
  const open = attributes.length ? `<td ${attributes.join(' ')}>` : '<td>'
  return `${open}${htmlText(text)}</td>`
}

/**
 * Merges intersected with the selection: `starts` maps the visible top-left of every merged region
 * to its spans and real anchor, `covered` holds the cells the table must not repeat. A merge that
 * starts before the selection is clipped to the selection, so copying part of it still yields the
 * merged value once instead of an empty rectangle.
 */
function mergePlan(sheet, box) {
  const starts = new Map()
  const covered = new Set()
  for (const merge of sheet?.merges ?? []) {
    const r1 = Math.max(merge.r1, box.r1)
    const c1 = Math.max(merge.c1, box.c1)
    const r2 = Math.min(merge.r2, box.r2)
    const c2 = Math.min(merge.c2, box.c2)
    if (r2 < r1 || c2 < c1) continue

    const key = `${r1},${c1}`
    if (starts.has(key) || covered.has(key)) continue
    const rowspan = r2 - r1 + 1
    const colspan = c2 - c1 + 1
    if (rowspan > MAX_HTML_ROWSPAN || colspan > MAX_HTML_COLSPAN) {
      // html consumers clamp larger spans, which would shift every following cell; refuse instead
      // of pasting a table with broken geometry.
      throw new RangeError(
        `merged region needs rowspan ${rowspan} / colspan ${colspan}; ` +
        `HTML allows at most ${MAX_HTML_ROWSPAN} / ${MAX_HTML_COLSPAN}`
      )
    }
    starts.set(key, { rowspan, colspan, sourceR: merge.r1, sourceC: merge.c1 })
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (r === r1 && c === c1) continue
        covered.add(`${r},${c}`)
      }
    }
  }
  return { starts, covered }
}

/**
 * A selection as an HTML table for the clipboard: inline CSS from the model's resolved styles,
 * escaped text and Excel's line breaks, one cell per value, and merged regions as `colspan` /
 * `rowspan` with the covered cells left out. Plain-text formats are unchanged; a caller puts this
 * in `text/html` and keeps `selectionToTsv` as `text/plain`.
 *
 * Throws `RangeError` above `MAX_HTML_CELLS`, above `maxBytes` of serialized UTF-8, or when a
 * merged span exceeds what HTML can represent. Explicit refusal is better than a clipped table
 * that looks complete.
 */
export function selectionToHtml(sheet, selection, { formulas = false, maxBytes = MAX_HTML_BYTES } = {}) {
  const box = bounds(selection)
  if (!box || !sheet) return ''

  const cellCount = (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1)
  if (cellCount > MAX_HTML_CELLS) {
    throw new RangeError(`selection is too large for HTML clipboard (${cellCount} > ${MAX_HTML_CELLS} cells)`)
  }

  const byteLimit = Number(maxBytes) > 0 && Number.isFinite(Number(maxBytes))
    ? Math.floor(Number(maxBytes))
    : MAX_HTML_BYTES
  const open = '<table border="0" cellspacing="0" cellpadding="0" style="border-collapse:collapse"><tbody>'
  const close = '</tbody></table>'
  let bytes = utf8Length(open) + utf8Length(close)

  const { starts, covered } = mergePlan(sheet, box)
  const rows = []
  for (let r = box.r1; r <= box.r2; r++) {
    bytes += ROW_TAG_BYTES
    if (bytes > byteLimit) throw new RangeError(`selection HTML exceeds ${byteLimit} bytes`)
    const cells = []
    for (let c = box.c1; c <= box.c2; c++) {
      const key = `${r},${c}`
      if (covered.has(key)) continue
      const span = starts.get(key)
      const cell = span
        ? htmlCell(cellAt(sheet, span.sourceR, span.sourceC), span, formulas)
        : htmlCell(cellAt(sheet, r, c), null, formulas)
      const size = utf8Length(cell)
      if (bytes + size > byteLimit) throw new RangeError(`selection HTML exceeds ${byteLimit} bytes`)
      bytes += size
      cells.push(cell)
    }
    rows.push(`<tr>${cells.join('')}</tr>`)
  }

  return `${open}${rows.join('')}${close}`
}

/* ------------------------------------------------------------------ copy ------------------------ */

/**
 * Copies through the one path the isolated frame allows: select a scratch textarea and run the
 * legacy `copy` command, which fires a `copy` event we can fill. Returns whether the copy
 * happened and leaves no node or focus change behind.
 */
function copyThrough(data, target) {
  if (!target?.body || typeof target.execCommand !== 'function') return false
  const previous = target.activeElement
  const area = target.createElement('textarea')
  area.value = data.find(([type]) => type === 'text/plain')?.[1] ?? ''
  area.setAttribute('readonly', '')
  area.setAttribute('aria-hidden', 'true')
  area.style.position = 'fixed'
  area.style.top = '0'
  area.style.left = '-9999px'
  area.style.opacity = '0'
  const onCopy = (event) => {
    if (event.clipboardData?.setData) {
      for (const [type, value] of data) event.clipboardData.setData(type, value)
      event.preventDefault()
    }
  }
  area.addEventListener('copy', onCopy)
  target.body.append(area)
  area.select()
  let copied = false
  try {
    copied = target.execCommand('copy') === true
  } catch {
    copied = false
  } finally {
    area.removeEventListener('copy', onCopy)
    area.remove()
    if (previous && previous !== area && typeof previous.focus === 'function') {
      try {
        previous.focus()
      } catch {
        // The previous element may be gone; focus restoration is best effort.
      }
    }
  }
  return copied
}

export function copyText(text, target = globalThis.document) {
  return copyThrough([['text/plain', String(text)]], target)
}

/**
 * Copies HTML with a plain-text fallback. `text` should be the matching TSV, so a target that
 * ignores `text/html` still receives the same rectangle. Returns whether the copy happened.
 */
export function copyHtml(html, text = '', target = globalThis.document) {
  return copyThrough([
    ['text/plain', String(text)],
    ['text/html', String(html)]
  ], target)
}
