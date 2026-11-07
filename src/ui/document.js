// The only file that opens an Orca document session. It keeps the zip of the last accepted file in
// step with what the host stores, and turns the edit log into the bytes the host saves.
//
// Why encode rebuilds from module state instead of its argument: `session` calls `encode(value)`
// synchronously on every send, and the value it passes is whatever `change()` last received. The
// file to produce is the last accepted file plus the edits made since that save mark, so that is
// what we rebuild; `decode` tells us when the host accepted it.

import { readZip, writeZip } from '../write/zip.js'
import { applyEdits } from '../write/xlsx-writer.js'
import { applyStructural } from '../write/structural.js'
import { loadWorkbook } from '../model/workbook.js'
import { createEditLog } from '../edit/edits.js'

let pkg = null // zip entries of the last accepted file
let acceptedBytes = null // bytes of the last accepted file
let pending = null // what the last encode produced, until decode confirms the host took it
let saveMark = null // where the edit log stood at the last accepted write
let date1904 = false // the workbook's date system, needed to write a Date back as a serial
const log = createEditLog()
let onRefused = () => {}
let onAccepted = () => {}
let onDiscarded = () => {}

export function subscribeDocument(handlers = {}) {
  if (handlers.refused) onRefused = handlers.refused
  if (handlers.accepted) onAccepted = handlers.accepted
  if (handlers.discarded) onDiscarded = handlers.discarded
}

/** Called after the workbook model loads: dates must be written in the file's own date system. */
export function setDateSystem(is1904) {
  date1904 = Boolean(is1904)
}

function slice(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
}

function sameBytes(a, b) {
  if (a === b) return true
  if (!a || !b || a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

/** Excel serial day number for a Date, in the workbook's own 1900 or 1904 system. */
function toSerial(value) {
  const date = value instanceof Date ? value : new Date(value)
  const utc = Date.UTC(
    date.getUTCFullYear(),
    date.getUTCMonth(),
    date.getUTCDate(),
    date.getUTCHours(),
    date.getUTCMinutes(),
    date.getUTCSeconds()
  )
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  return (utc - epoch) / 86_400_000
}

/**
 * The cached result of a formula, or null when the file carries none. A formula whose result the
 * model could not resolve holds the formula text in `value`; that must never be written as a
 * cached result, or the next reader shows a string where a number belongs.
 */
function cachedResult(value) {
  if (value.kind === 'number' && Number.isFinite(value.value)) return Number(value.value)
  if (value.kind === 'boolean') return Boolean(value.value)
  if (value.kind === 'date') return toSerial(value.value)
  return null
}

/** Edit-log operation to the writer's cell edit: `before` is not the writer's business. */
function toCellEdit(op) {
  const value = op.after ?? { kind: 'blank' }
  if (value.formula) {
    return {
      sheet: op.sheet,
      ref: op.ref,
      value: { kind: 'formula', formula: value.formula, cached: cachedResult(value) }
    }
  }
  switch (value.kind) {
    case 'number':
      return { sheet: op.sheet, ref: op.ref, value: { kind: 'number', value: Number(value.value) } }
    case 'date':
      return { sheet: op.sheet, ref: op.ref, value: { kind: 'date', value: toSerial(value.value) } }
    case 'boolean':
      return { sheet: op.sheet, ref: op.ref, value: { kind: 'boolean', value: Boolean(value.value) } }
    case 'error':
      return { sheet: op.sheet, ref: op.ref, value: { kind: 'error', value: String(value.value) } }
    case 'blank':
      return { sheet: op.sheet, ref: op.ref, value: { kind: 'blank' } }
    default:
      return { sheet: op.sheet, ref: op.ref, value: { kind: 'text', value: String(value.value ?? '') } }
  }
}

function hasUnsaved() {
  if (pending) return true
  if (!saveMark) return log.ops().length > 0
  return log.opsSince(saveMark).length > 0
}

// Apply cells in batches separated by structural snapshots. A snapshot represents
// the entire package at that point in history, including cross-sheet references.
function materialize(entries, ops) {
  let current = entries
  let cells = []
  function flush() {
    if (!cells.length) return null
    const result = applyEdits(current, cells.map(toCellEdit))
    cells = []
    if (result.status !== 'ok') return result
    current = result.entries
    return null
  }
  for (const op of ops) {
    if (op.type === 'cell') {
      cells.push(op)
      continue
    }
    const failure = flush()
    if (failure) return failure
    current = op.type === 'restore' ? op.entries : op.afterEntries
  }
  const failure = flush()
  return failure ?? { status: 'ok', entries: current }
}

function currentEntries() {
  if (!pkg) return { status: 'refused', reason: 'this file could not be read, so changes cannot be written back' }
  if (blocked.size) return { status: 'refused', reason: 'a previous edit was refused; reload the document before changing its structure' }
  return materialize(pkg, saveMark ? log.opsSince(saveMark) : log.ops())
}

/** Bytes of the currently staged workbook, for rebuilding the view after structural history. */
export function currentDocumentBytes() {
  const result = currentEntries()
  if (result.status !== 'ok') throw new Error(result.reason)
  return writeZip(result.entries)
}

function decode(content) {
  if (content.kind !== 'bytes') throw new TypeError('a workbook editor needs a bytes document')
  const bytes = new Uint8Array(content.bytes)

  if (pending && sameBytes(bytes, pending.bytes)) {
    // The host accepted the edit this tab sent; the file on disk is exactly these bytes now.
    pkg = pending.entries
    acceptedBytes = bytes
    saveMark = pending.mark
    pending = null
    onAccepted()
    return { bytes }
  }

  // A document we did not write: first open, an outside change, or a reload after a conflict.
  // Anything the log still holds describes edits relative to a file that no longer exists, so it
  // is dropped rather than replayed onto someone else's bytes.
  const discarded = hasUnsaved()
  blocked.clear()
  acceptedBytes = bytes
  const zip = readZip(bytes)
  pkg = zip.status === 'ok' ? zip.entries : null
  pending = null
  log.clear()
  saveMark = log.markSaved()
  if (discarded) onDiscarded()
  return { bytes, zipStatus: zip.status }
}

function encode() {
  if (!acceptedBytes) throw new Error('the workbook has not been read yet')
  if (pending) throw new Error('a workbook save is already awaiting host acceptance')
  if (!pkg) {
    // A file this editor cannot take apart must never look as if it were saved.
    const reason = 'this file could not be read, so changes cannot be written back'
    onRefused(reason)
    throw new Error(reason)
  }
  const ops = saveMark ? log.opsSince(saveMark) : log.ops()
  if (!ops.length) {
    // Nothing to write. This is still an answer to our own change(), so the host's echo must be
    // recognised as ours: otherwise decode treats it as an outside write and wipes the undo history.
    pending = { bytes: acceptedBytes, entries: pkg, mark: log.markSaved() }
    return { kind: 'bytes', bytes: slice(acceptedBytes) }
  }

  const structural = ops.some((op) => op.type !== 'cell')
  const usable = ops.filter((op) => op.type !== 'cell' || !blocked.has(cellKey(op)))
  if (structural && blocked.size) {
    const reason = 'a refused cell edit prevents saving the structural transaction'
    onRefused(reason)
    throw new Error(reason)
  }
  const result = materialize(pkg, usable)
  if (result.status !== 'ok') {
    if (structural) {
      onRefused(result.reason)
      throw new Error(result.reason)
    }
    // Find which edits the writer will not take, remember them, and save the rest. A single cell
    // inside a merged range must not cost the person every other edit in the tab.
    const offender = firstRefused(pkg, usable)
    if (offender) {
      blocked.add(cellKey(offender))
      onRefused(`cell ${offender.ref} on sheet "${offender.sheet}" cannot be changed here`)
      const rest = usable.filter((op) => op !== offender)
      const retry = rest.length ? applyEdits(pkg, rest.map(toCellEdit)) : { status: 'ok', entries: pkg }
      if (retry.status !== 'ok') {
        onRefused(retry.reason)
        throw new Error(retry.reason)
      }
      const bytes = writeZip(retry.entries)
      pending = { bytes, entries: retry.entries, mark: log.markSaved() }
      return { kind: 'bytes', bytes: slice(bytes) }
    }
    onRefused(result.reason)
    throw new Error(result.reason)
  }
  const bytes = writeZip(result.entries)
  pending = { bytes, entries: result.entries, mark: log.markSaved() }
  return { kind: 'bytes', bytes: slice(bytes) }
}

/** The first edit the writer refuses on its own, found by bisecting the batch. */
function firstRefused(entries, ops) {
  let low = 0
  let high = ops.length - 1
  let found = null
  while (low <= high) {
    const middle = (low + high) >> 1
    const attempt = applyEdits(entries, ops.slice(0, middle + 1).map(toCellEdit))
    if (attempt.status === 'ok') {
      low = middle + 1
    } else {
      found = ops[middle]
      high = middle - 1
    }
  }
  return found
}

export async function openDocument() {
  const api = globalThis.orca
  if (!api?.document) throw new Error('this file runs as a custom editor')
  return api.document.session({ decode, encode })
}

/**
 * Records one cell edit and asks the session to save it after its debounce.
 *
 * Returns `{ ok: true, op }` with the settled operation, or `{ ok: false, reason }`.
 *
 * Expensive validation (running the writer) is deliberately not done here: it cost a full pass over
 * the sheet on every keystroke. `encode` filters out anything the writer refuses instead, so one bad
 * edit can never block the rest of the save.
 */
export function commitEdit(doc, op) {
  const applied = log.apply(op)
  if (!applied) return { ok: true, op: { ...op, after: op.before } }
  doc.change({})
  return { ok: true, op: applied }
}

/** Preflight the complete staged package before recording a structural change. */
export function commitStructural(doc, operation) {
  if (doc.readOnly) return { ok: false, reason: 'the document is read-only' }
  if (pending) return { ok: false, reason: 'a workbook save is awaiting host acceptance' }
  const before = currentEntries()
  if (before.status !== 'ok') return { ok: false, reason: before.reason }
  const result = applyStructural(before.entries, operation)
  if (result.status !== 'ok') return { ok: false, reason: result.reason }
  const state = log.checkpoint()
  try {
    const op = log.apply({ ...operation, type: 'structural', beforeEntries: before.entries, afterEntries: result.entries })
    const bytes = writeZip(result.entries)
    doc.change({})
    return { ok: true, op, bytes }
  } catch (error) {
    log.restore(state)
    return { ok: false, reason: error.message }
  }
}

/** Marks a cell the writer refused, so later saves skip it instead of failing whole. */
const blocked = new Set()
export function isBlocked() {
  return blocked.size > 0
}
const cellKey = (op) => `${op.sheet}\u0000${op.ref}`

function moveHistory(doc, direction) {
  const next = direction === 'undo' ? log.peekUndo() : log.peekRedo()
  if (!next) return null
  if (doc.readOnly || (pending && next.type === 'structural')) {
    onRefused(doc.readOnly ? 'the document is read-only' : 'a workbook save is awaiting host acceptance')
    return null
  }
  const state = log.checkpoint()
  try {
    const op = log[direction]()
    doc.change({})
    return op
  } catch (error) {
    log.restore(state)
    onRefused(error.message)
    return null
  }
}

export function undo(doc) {
  return moveHistory(doc, 'undo')
}

export function redo(doc) {
  return moveHistory(doc, 'redo')
}

export function editState() {
  return { canUndo: log.canUndo, canRedo: log.canRedo, dirty: hasUnsaved() }
}

export async function loadModel(bytes) {
  return loadWorkbook(bytes)
}
