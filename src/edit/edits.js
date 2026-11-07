/**
 * Edit log for the workbook: cell operations, undo/redo and the save delta.
 *
 * Pure ESM. No DOM, no Orca API. Structural transactions retain package snapshots
 * so deletion can be undone without reconstructing lost cells or metadata.
 *
 * @typedef {{ kind: 'number' | 'text' | 'boolean' | 'date' | 'error' | 'blank',
 *   value: unknown, display: string, formula: string | null }} CellValue
 *
 * @typedef {{ type: 'cell', sheet: string, ref: string,
 *   before: CellValue | null, after: CellValue | null }} CellOp
 *
 * @typedef {{ seq: number }} SaveMark
 *   Opaque mark from markSaved(); `seq` is monotone for the lifetime of the log,
 *   so a mark taken before clear() still means "everything after that point".
 */

/** Max 0-based column index: XFD. */
export const EXCEL_MAX_COL = 16383
/** Max 0-based row index: 1048576. */
export const EXCEL_MAX_ROW = 1048575
export const DEFAULT_UNDO_LIMIT = 1000
export const DEFAULT_COALESCE_MS = 600

const CELL_KINDS = new Set(['number', 'text', 'boolean', 'date', 'error', 'blank'])
const A1 = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/
const DEEP_LIMIT = 8

/* ---------------------------------------------------------------- addresses */

/** 0 -> 'A', 26 -> 'AA', 16383 -> 'XFD'. */
export function columnName(index) {
  if (!Number.isInteger(index) || index < 0 || index > EXCEL_MAX_COL) {
    throw new RangeError(`columnName: column index out of range: ${String(index)}`)
  }
  let n = index + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** 'A' -> 0, 'AA' -> 26, 'XFD' -> 16383. */
export function columnIndex(letters) {
  if (typeof letters !== 'string' || !/^[A-Za-z]{1,3}$/.test(letters)) {
    throw new TypeError(`columnIndex: expected 1..3 letters, got ${String(letters)}`)
  }
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  const index = n - 1
  if (index > EXCEL_MAX_COL) throw new RangeError(`columnIndex: column out of range: ${letters}`)
  return index
}

/** 'B2' -> { r: 1, c: 1 }. Accepts $B$2, any case. */
export function parseRef(ref) {
  if (typeof ref !== 'string') throw new TypeError(`parseRef: expected a string, got ${typeof ref}`)
  const m = A1.exec(ref.trim())
  if (!m) throw new TypeError(`parseRef: invalid A1 reference: ${ref}`)
  const c = columnIndex(m[1])
  const r = Number(m[2]) - 1
  if (!Number.isInteger(r) || r < 0 || r > EXCEL_MAX_ROW) {
    throw new RangeError(`parseRef: row out of range: ${ref}`)
  }
  return { r, c }
}

/** { r: 1, c: 1 } -> 'B2'. */
export function formatRef(pos) {
  if (!pos || typeof pos !== 'object') throw new TypeError('formatRef: expected { r, c }')
  const { r, c } = pos
  if (!Number.isInteger(r) || r < 0 || r > EXCEL_MAX_ROW) {
    throw new RangeError(`formatRef: row out of range: ${String(r)}`)
  }
  if (!Number.isInteger(c) || c < 0 || c > EXCEL_MAX_COL) {
    throw new RangeError(`formatRef: column out of range: ${String(c)}`)
  }
  return `${columnName(c)}${r + 1}`
}

/** 'B2', 1, 1 -> 'C3'. Throws RangeError when the result leaves the sheet. */
export function shiftRef(ref, dr, dc) {
  if (!Number.isInteger(dr) || !Number.isInteger(dc)) {
    throw new TypeError('shiftRef: dr and dc must be integers')
  }
  const { r, c } = parseRef(ref)
  return formatRef({ r: r + dr, c: c + dc })
}

/* ------------------------------------------------------------------- values */

/**
 * Structural equality for raw CellValue payloads (Date, arrays, plain objects).
 * Null and undefined are equal: a missing cached value is a missing cached value.
 */
function deepValueEqual(a, b, depth) {
  if (Object.is(a, b)) return true
  if (a == null || b == null) return a == null && b == null
  if (typeof a !== typeof b) return false
  if (typeof a !== 'object') return false
  if (depth >= DEEP_LIMIT) return false
  if (a instanceof Date || b instanceof Date) {
    return a instanceof Date && b instanceof Date && a.getTime() === b.getTime()
  }
  if (Array.isArray(a) !== Array.isArray(b)) return false
  const ka = Object.keys(a)
  const kb = Object.keys(b)
  if (ka.length !== kb.length) return false
  for (const key of ka) {
    if (!Object.prototype.hasOwnProperty.call(b, key)) return false
    if (!deepValueEqual(a[key], b[key], depth + 1)) return false
  }
  return true
}

/**
 * CellValue equality by value. `display` is derived from kind/value/numFmt and is
 * ignored; kind, value and formula are compared. Used to drop no-op edits.
 */
export function equals(a, b) {
  if (a === b) return true
  if (a == null || b == null) return a == null && b == null
  if (typeof a !== 'object' || typeof b !== 'object') return false
  if (a.kind !== b.kind) return false
  if ((a.formula ?? null) !== (b.formula ?? null)) return false
  return deepValueEqual(a.value, b.value, 0)
}

/* ---------------------------------------------------------------------- log */

function assertCellValue(value, label) {
  if (value == null) return null
  if (typeof value !== 'object') throw new TypeError(`${label} must be a CellValue or null`)
  if (!CELL_KINDS.has(value.kind)) {
    throw new TypeError(`${label}.kind must be one of ${[...CELL_KINDS].join(', ')}; got ${String(value.kind)}`)
  }
  return value
}

function normalizeOp(op) {
  if (!op || typeof op !== 'object') throw new TypeError('apply: op must be an object')
  if (op.type === 'structural') {
    if (typeof op.sheet !== 'string' || !op.sheet ||
        !['row', 'column'].includes(op.axis) || !['insert', 'delete'].includes(op.action) ||
        !Number.isInteger(op.index) || !Number.isInteger(op.count) ||
        op.index < 1 || op.count < 1 ||
        op.index + op.count - 1 > (op.axis === 'row' ? EXCEL_MAX_ROW + 1 : EXCEL_MAX_COL + 1) ||
        !Array.isArray(op.beforeEntries) || !Array.isArray(op.afterEntries)) {
      throw new TypeError('apply: invalid structural transaction')
    }
    return { ...op }
  }
  if (op.type !== 'cell') throw new TypeError(`apply: unsupported op type: ${String(op.type)}`)
  if (typeof op.sheet !== 'string' || op.sheet === '') {
    throw new TypeError('apply: op.sheet must be a non-empty string')
  }
  parseRef(op.ref) // validates the address
  return {
    type: 'cell',
    sheet: op.sheet,
    ref: op.ref,
    before: assertCellValue(op.before, 'op.before'),
    after: assertCellValue(op.after, 'op.after')
  }
}

/** Identity of a cell for coalescing and for the save delta. */
function cellKey(op) {
  return `${op.sheet}\u0000${formatRef(parseRef(op.ref))}`
}

function markSeq(mark) {
  if (typeof mark === 'number' && Number.isInteger(mark) && mark >= 0) return mark
  if (mark && typeof mark === 'object' && Number.isInteger(mark.seq)) return mark.seq
  throw new TypeError('opsSince: pass the value returned by markSaved()')
}

/**
 * @param {{ coalesceWindowMs?: number, undoLimit?: number,
 *   coalesce?: boolean, now?: () => number }} [options]
 * @returns {{
 *   apply: (op: CellOp) => CellOp | null,
 *   undo: () => CellOp | null,
 *   redo: () => CellOp | null,
 *   ops: () => CellOp[],
 *   opsSince: (mark: SaveMark | number) => CellOp[],
 *   markSaved: () => SaveMark,
 *   clear: () => void,
 *   readonly canUndo: boolean,
 *   readonly canRedo: boolean,
 *   readonly undoSize: number,
 *   readonly dropped: number
 * }}
 *
 * `ops()` is the full application order and keeps growing: it is the writer input.
 * The undo stack is only the retained tail (undoLimit, default 1000); `dropped`
 * counts operations that fell out of it and can no longer be undone.
 */
export function createEditLog(options = {}) {
  const coalesceWindowMs = options.coalesceWindowMs ?? DEFAULT_COALESCE_MS
  const undoLimit = options.undoLimit ?? DEFAULT_UNDO_LIMIT
  const coalesce = options.coalesce !== false
  const now = options.now ?? Date.now

  if (!Number.isFinite(coalesceWindowMs) || coalesceWindowMs < 0) {
    throw new RangeError('createEditLog: coalesceWindowMs must be a non-negative number')
  }
  if (!Number.isInteger(undoLimit) || undoLimit < 1) {
    throw new RangeError('createEditLog: undoLimit must be a positive integer')
  }

  /** Monotone for the lifetime of the log, so old marks stay meaningful after clear(). */
  let seq = 0
  /** @type {{ seq: number, at: number, key: string, op: CellOp }[]} applied operations, in order */
  let log = []
  /** @type {typeof log} undo window: the last <= undoLimit records of `log` */
  let undoStack = []
  /** @type {typeof log} undone records, last undone on top */
  let redoStack = []
  let dropped = 0
  /**
   * Highest seq that has been handed to the writer by markSaved(). Records at or
   * below it may already be in the file, so undoing one has to be written back as
   * a new operation instead of vanishing from the log.
   */
  let savedSeq = 0
  /** True while the tail of `log` may still absorb an apply as a coalesced keystroke. */
  let chainOpen = false

  function trimUndo() {
    while (undoStack.length > undoLimit) {
      undoStack.shift()
      dropped++
    }
  }

  function apply(op) {
    const next = normalizeOp(op)
    if (next.type === 'cell' && equals(next.before, next.after)) return null

    redoStack.length = 0
    const at = now()
    const key = next.type === 'cell' ? cellKey(next) : null
    const last = log.length > 0 ? log[log.length - 1] : null

    // A save closes the batch: the value at the mark is already in the file, so the
    // next keystroke must start a new operation even inside the time window.
    if (next.type === 'cell' && coalesce && chainOpen && last && last.key === key && at - last.at <= coalesceWindowMs) {
      // Fast typing in one cell: keep the first `before`, take the newest `after`.
      const merged = {
        type: 'cell',
        sheet: next.sheet,
        ref: next.ref,
        before: last.op.before,
        after: next.after
      }
      if (equals(merged.before, merged.after)) {
        // Typed back to the original content: the whole cell edit never happened. The caller still
        // needs to know what the cell shows now, so the effective operation is returned with
        // `after` equal to `before` instead of a bare null.
        log.pop()
        undoStack.pop()
        return { ...merged, after: merged.before, collapsed: true }
      }
      last.op = merged
      last.at = at
      return merged
    }

    const record = { seq: ++seq, at, key, op: next }
    log.push(record)
    undoStack.push(record)
    trimUndo()
    chainOpen = next.type === 'cell'
    return next
  }

  /**
   * Undoes the last applied operation and returns it, or null when there is none.
   *
   * Above the save boundary (the writer never saw the record) it is dropped from
   * the log: the file is already correct without it. At or below the boundary the
   * record may already be on disk, so the reversion is appended as an inverse
   * operation and reaches the writer as a real delta. `before` of the returned op
   * still carries the pre-edit value, so the caller can restore the cell in the view.
   */
  function undo() {
    const record = undoStack[undoStack.length - 1]
    if (!record) return null
    undoStack.pop()
    chainOpen = false

    if (record.seq > savedSeq) {
      if (log[log.length - 1] !== record) {
        throw new Error('edit log: undo stack out of sync with the operation log')
      }
      log.pop()
    } else {
      log.push({
        seq: ++seq,
        at: now(),
        key: record.key,
        op: record.op.type === 'structural'
          ? { type: 'restore', entries: record.op.beforeEntries }
          : {
              type: 'cell',
              sheet: record.op.sheet,
              ref: record.op.ref,
              before: record.op.after,
              after: record.op.before
            }
      })
    }

    redoStack.push(record)
    return record.op
  }

  /** Re-applies the last undone operation as a new log entry and returns it. */
  function redo() {
    const record = redoStack[redoStack.length - 1]
    if (!record) return null
    redoStack.pop()

    // A fresh seq keeps the re-applied op inside every mark taken before it.
    const replay = { ...record, seq: ++seq, at: now() }
    log.push(replay)
    undoStack.push(replay)
    trimUndo()
    chainOpen = false
    return record.op
  }

  /**
   * Everything the writer has applied so far, in application order. Undo and redo
   * append here too: a reversion that crossed a save boundary is a real operation.
   */
  function ops() {
    return log.map((record) => record.op)
  }

  /**
   * What the writer has to flush: for every cell touched after `mark`, one op with
   * `before` at the mark and the newest `after`. Cells changed back to their
   * mark-time value are omitted, so a flush is a real delta.
   *
   * Dropping such a cell is only sound inside one batch: the first post-mark `before`
   * is the value the file has at the mark. A save starts a new batch and therefore a
   * new mark, so an undo that rewinds a saved edit keeps its own `before` (the saved
   * value) and survives this filter.
   */
  function opsSince(mark) {
    const since = markSeq(mark)
    const out = []
    const touched = new Map()
    function flush() {
      for (const { first, last } of touched.values()) {
        if (equals(first.before, last.after)) continue
        out.push({ type: 'cell', sheet: last.sheet, ref: last.ref, before: first.before, after: last.after })
      }
      touched.clear()
    }
    for (const record of log) {
      if (record.seq <= since) continue
      const op = record.op
      if (op.type !== 'cell') {
        flush()
        out.push(op)
      } else {
        const seen = touched.get(record.key)
        if (seen) seen.last = op
        else touched.set(record.key, { first: op, last: op })
      }
    }
    flush()
    return out
  }

  /**
   * Records that everything applied so far is in the file and returns the mark
   * `opsSince` understands. Marks the batch boundary: the next apply is a new
   * operation, never a coalesced continuation of the saved one.
   */
  function markSaved() {
    savedSeq = seq
    chainOpen = false
    return Object.freeze({ seq })
  }

  // Restore the entire history if the host rejects a change synchronously.
  function checkpoint() {
    const copies = new Map()
    const copy = (record) => {
      if (!copies.has(record)) copies.set(record, { ...record })
      return copies.get(record)
    }
    return {
      seq, savedSeq, dropped, chainOpen,
      log: log.map(copy), undoStack: undoStack.map(copy), redoStack: redoStack.map(copy)
    }
  }

  function restore(state) {
    ;({ seq, savedSeq, dropped, chainOpen, log, undoStack, redoStack } = state)
  }

  function clear() {
    log = []
    undoStack = []
    redoStack = []
    dropped = 0
    chainOpen = false
    savedSeq = seq
  }

  return {
    apply,
    undo,
    redo,
    ops,
    opsSince,
    markSaved,
    clear,
    checkpoint,
    restore,
    peekUndo: () => undoStack.at(-1)?.op ?? null,
    peekRedo: () => redoStack.at(-1)?.op ?? null,
    get canUndo() {
      return undoStack.length > 0
    },
    get canRedo() {
      return redoStack.length > 0
    },
    get undoSize() {
      return undoStack.length
    },
    get dropped() {
      return dropped
    }
  }
}
