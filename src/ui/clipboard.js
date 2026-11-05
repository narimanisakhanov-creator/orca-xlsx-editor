// Clipboard text for the grid: TSV, CSV and Markdown built from a SheetModel.
// The formatters are pure functions (no DOM, no Orca API) so `node --test` runs them directly;
// only copyText touches the DOM, because the editor frame has no navigator.clipboard.

const TAB_OR_NEWLINE = /[\t\r\n]+/g

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

/**
 * Copies through the one path the isolated frame allows: select a scratch textarea and run the
 * legacy `copy` command, which fires a `copy` event we can fill. Returns whether the copy
 * happened and leaves no node or focus change behind.
 */
export function copyText(text, target = globalThis.document) {
  if (!target?.body || typeof target.execCommand !== 'function') return false
  const previous = target.activeElement
  const area = target.createElement('textarea')
  area.value = String(text)
  area.setAttribute('readonly', '')
  area.setAttribute('aria-hidden', 'true')
  area.style.position = 'fixed'
  area.style.top = '0'
  area.style.left = '-9999px'
  area.style.opacity = '0'
  const onCopy = (event) => {
    if (event.clipboardData?.setData) {
      event.clipboardData.setData('text/plain', area.value)
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
