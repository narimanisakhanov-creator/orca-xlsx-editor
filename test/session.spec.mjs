// The session rules that keep the file and the grid honest. Each test covers a way an edit could be
// lost or silently mislabelled saved; every one of them failed before the fixes.
import { test } from 'node:test'
import assert from 'node:assert/strict'

import { readZip, writeZip } from '../src/write/zip.js'
import ExcelJS from 'exceljs'
import { fixtureBytes } from './helpers/fixtures.mjs'

const copy = (bytes) => Uint8Array.from(bytes)

/** The parts of a workbook as name to bytes, through the plugin's own reader. */
function parts(bytes) {
  const result = readZip(bytes)
  assert.equal(result.status, 'ok', `unreadable zip: ${result.status}`)
  return Object.fromEntries(result.entries.map((entry) => [entry.name, entry.data]))
}

const sheetText = (bytes, part = 'xl/worksheets/sheet1.xml') =>
  Buffer.from(parts(bytes)[part]).toString('utf8')

/** A workbook whose one worksheet part is rewritten, with every other part left as it was. */
function withSheet(bytes, part, transform) {
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok')
  const entries = zip.entries.map((entry) => {
    if (entry.name !== part) return entry
    const next = new TextEncoder().encode(transform(Buffer.from(entry.data).toString('utf8')))
    return { ...entry, data: next, changed: true }
  })
  return writeZip(entries)
}

/**
 * Stands in for Orca's document session: one edit in flight, a version per accepted write, an
 * outside write that can land at any time.
 */
function createHost(bytes) {
  const state = { version: 1, bytes: copy(bytes), readOnly: false, discarded: 0, failChange: false, revision: 0 }
  let handlers = null
  let timer = null
  let inFlight = null
  let dirty = false
  const replaceHandlers = new Set()

  const content = (buffer) => ({
    version: state.version,
    readOnly: state.readOnly,
    content: {
      kind: 'bytes',
      bytes: buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
    }
  })

  async function send() {
    const revision = state.revision
    const encoded = handlers.encode()
    await Promise.resolve()
    state.bytes = new Uint8Array(encoded.bytes)
    state.version += 1
    dirty = state.revision !== revision
    handlers.decode(encoded)
  }

  const api = {
    document: {
      session: async (options) => {
        handlers = options
        handlers.decode(content(state.bytes).content)
        return {
          get value() {
            return handlers.decode(content(state.bytes).content)
          },
          get readOnly() {
            return state.readOnly
          },
          change() {
            if (state.failChange) throw new Error('host rejected change')
            if (state.readOnly) return
            state.revision++
            clearTimeout(timer)
            timer = setTimeout(() => {
              timer = null
              if (!inFlight) inFlight = send().finally(() => (inFlight = null))
            }, 0)
            dirty = true
          },
          async flush() {
            clearTimeout(timer)
            timer = null
            if (inFlight) return inFlight
            if (!dirty) return undefined
            inFlight = send().finally(() => (inFlight = null))
            return inFlight
          },
          onReplace(fn) {
            replaceHandlers.add(fn)
            return { dispose: () => replaceHandlers.delete(fn) }
          }
        }
      }
    }
  }

  return {
    api,
    setReadOnly(value) { state.readOnly = value },
    setFailChange(value) { state.failChange = value },
    bytes: () => copy(state.bytes),
    /** An outside writer replaced the file, and Orca handed the tab the new document. */
    outsideWrite(next) {
      state.bytes = copy(next)
      state.version += 1
      const delivered = handlers.decode(content(state.bytes).content)
      Array.from(replaceHandlers).forEach((fn) => fn(delivered))
    }
  }
}

async function withHost(file, body) {
  const original = fixtureBytes(file)
  const host = createHost(original)
  globalThis.orca = host.api
  try {
    const session = await import(`../src/ui/document.js?case=${Math.random()}`)
    return await body(session, host, original)
  } finally {
    delete globalThis.orca
  }
}

const number = (value) => ({ kind: 'number', value, display: String(value), formula: null })
const text = (value) => ({ kind: 'text', value, display: value, formula: null })
const settle = () => new Promise((resolve) => setTimeout(resolve, 10))

test('saving with no edits returns the accepted bytes untouched', async () => {
  await withHost('rich.xlsx', async (session, host, original) => {
    const doc = await session.openDocument()
    await doc.flush()
    assert.deepEqual(Buffer.from(host.bytes()), Buffer.from(original))
    assert.equal(session.editState().dirty, false)
  })
})

test('an edit the writer refuses does not poison the other edits in the tab', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    const refusals = []
    session.subscribeDocument({ refused: (reason) => refusals.push(reason) })

    // C9 sits inside the merged range A9:E9, which the writer will not touch.
    const blocked = session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'C9',
      before: text(''),
      after: number(1)
    })
    assert.equal(blocked.ok, true, 'commit is cheap; the writer decides at save time')

    const accepted = session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: number(282349),
      after: number(4242)
    })
    assert.equal(accepted.ok, true)
    await doc.flush()
    await settle()

    // The merges cell is refused, the normal edit is written, and the tab is clean again.
    assert.equal(refusals.length, 1)
    assert.match(refusals[0], /C9/)
    assert.match(sheetText(host.bytes()), /<c r="B2"[^>]*><v>4242<\/v>/)
    assert.equal(session.editState().dirty, false)
  })
})

test('an outside write clears the history so undo cannot resurrect a stale value', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: number(282349),
      after: number(100)
    })
    await doc.flush()
    await settle()

    const theirs = withSheet(host.bytes(), 'xl/worksheets/sheet1.xml', (xml) =>
      xml.replace(/<c r="B2"[^>]*><v>100<\/v><\/c>/, '<c r="B2" s="3"><v>777</v></c>')
    )
    host.outsideWrite(theirs)

    assert.equal(session.editState().canUndo, false, 'the log must not outlive the bytes it described')
    assert.equal(session.undo(doc), null)

    await doc.flush()
    await settle()
    assert.match(sheetText(host.bytes()), /<c r="B2"[^>]*><v>777<\/v>/)
  })
})

test('an outside write does not leave an unsaved edit looking saved', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: number(282349),
      after: number(42)
    })
    assert.equal(session.editState().dirty, true)

    host.outsideWrite(host.bytes())

    // The outside version won and the edit is gone; a clean tab that still shows the edit would be
    // the lie this test exists for.
    assert.equal(session.editState().dirty, false)
    assert.equal(session.editState().canUndo, false)

    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B3',
      before: number(271219),
      after: number(5)
    })
    await doc.flush()
    await settle()
    const sheet = sheetText(host.bytes())
    assert.match(sheet, /<c r="B3"[^>]*><v>5<\/v>/)
    assert.doesNotMatch(sheet, /<c r="B2"[^>]*><v>42<\/v>/)
  })
})

test('a flush with nothing to write is our own echo, not an outside write', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: number(282349),
      after: number(1)
    })
    await doc.flush()
    await settle()
    assert.equal(session.editState().canUndo, true)

    // Saving again with an empty delta used to look like a foreign document and wipe the history.
    await doc.flush()
    await settle()
    assert.equal(session.editState().canUndo, true, 'an empty save must not clear the history')
    assert.match(sheetText(host.bytes()), /<c r="B2"[^>]*><v>1<\/v>/)
  })
})

test('undo before the debounce keeps the rest of the history', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: number(282349),
      after: number(11)
    })
    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B3',
      before: number(271219),
      after: number(22)
    })
    // Undo lands before the debounce has sent anything.
    const undone = session.undo(doc)
    assert.equal(undone.ref, 'B3')

    await doc.flush()
    await settle()

    const sheet = sheetText(host.bytes())
    assert.match(sheet, /<c r="B2"[^>]*><v>11<\/v>/)
    assert.doesNotMatch(sheet, /<v>22<\/v>/)
    assert.equal(session.editState().canUndo, true, 'the remaining edit is still undoable')
    assert.equal(session.editState().canRedo, true, 'the undone edit is still redoable')
  })
})

test('a workbook the reader cannot take apart never reports a successful save', async () => {
  const garbage = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5])
  const host = createHost(garbage)
  globalThis.orca = host.api
  const refusals = []
  try {
    const session = await import(`../src/ui/document.js?case=${Math.random()}`)
    session.subscribeDocument({ refused: (reason) => refusals.push(reason) })
    const doc = await session.openDocument()
    session.commitEdit(doc, {
      type: 'cell',
      sheet: 'Sheet1',
      ref: 'A1',
      before: text(''),
      after: number(1)
    })

    await assert.rejects(() => doc.flush())
    assert.equal(refusals.length > 0, true)
    assert.equal(session.editState().dirty, true, 'an unwritable file must not look saved')
  } finally {
    delete globalThis.orca
  }
})

test('a deleted row and its metadata survive undo and redo across accepted saves', async () => {
  const book = new ExcelJS.Workbook()
  const sheet = book.addWorksheet('Data')
  sheet.getCell('A1').value = 1
  sheet.getCell('A2').value = 'removed'
  sheet.getCell('A2').font = { bold: true }
  sheet.getRow(2).height = 27
  sheet.getCell('A3').value = 3
  const host = createHost(new Uint8Array(await book.xlsx.writeBuffer()))
  globalThis.orca = host.api
  try {
    const session = await import(`../src/ui/document.js?case=${Math.random()}`)
    const doc = await session.openDocument()
    session.commitEdit(doc, { type: 'cell', sheet: 'Data', ref: 'A1', before: number(1), after: number(2) })
    const result = session.commitStructural(doc, { sheet: 'Data', axis: 'row', action: 'delete', index: 2, count: 1 })
    assert.equal(result.ok, true, result.reason)
    session.commitEdit(doc, { type: 'cell', sheet: 'Data', ref: 'A2', before: number(3), after: number(4) })
    await doc.flush()
    assert.match(sheetText(host.bytes()), /<c r="A2"[^>]*><v>4<\/v>/)
    session.undo(doc)
    await doc.flush()
    session.undo(doc)
    await doc.flush()
    const restored = new ExcelJS.Workbook()
    await restored.xlsx.load(host.bytes())
    assert.equal(restored.getWorksheet('Data').getCell('A2').value, 'removed')
    assert.equal(restored.getWorksheet('Data').getCell('A2').font.bold, true)
    assert.equal(restored.getWorksheet('Data').getRow(2).height, 27)
    assert.equal(restored.getWorksheet('Data').getCell('A1').value, 2)
    session.redo(doc)
    await doc.flush()
    session.redo(doc)
    await doc.flush()
    assert.match(sheetText(host.bytes()), /<c r="A2"[^>]*><v>4<\/v>/)
  } finally {
    delete globalThis.orca
  }
})

test('a structural refusal leaves the document and history untouched', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    const before = host.bytes()
    const state = session.editState()
    const result = session.commitStructural(doc, { sheet: 'missing', axis: 'row', action: 'delete', index: 1, count: 1 })
    assert.equal(result.ok, false)
    assert.deepEqual(session.editState(), state)
    await doc.flush()
    assert.deepEqual(host.bytes(), before)
  })
})

test('a host rejection and read-only state cannot create a structural history entry', async () => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Data').getCell('A1').value = 1
  const host = createHost(new Uint8Array(await book.xlsx.writeBuffer()))
  globalThis.orca = host.api
  try {
    const session = await import(`../src/ui/document.js?case=${Math.random()}`)
    const doc = await session.openDocument()
    const original = host.bytes()
    const op = { sheet: 'Data', axis: 'row', action: 'delete', index: 1, count: 1 }
    host.setReadOnly(true)
    assert.equal(session.commitStructural(doc, op).ok, false)
    host.setReadOnly(false)
    host.setFailChange(true)
    assert.equal(session.commitStructural(doc, op).ok, false)
    assert.deepEqual(session.editState(), { canUndo: false, canRedo: false, dirty: false })
    assert.deepEqual(host.bytes(), original)
  } finally {
    delete globalThis.orca
  }
})

test('edits made during an in-flight save are not marked accepted with its earlier bytes', async () => {
  await withHost('rich.xlsx', async (session, host) => {
    const doc = await session.openDocument()
    session.commitEdit(doc, { type: 'cell', sheet: 'Summary', ref: 'B2', before: number(282349), after: number(1) })
    const firstSave = doc.flush()
    session.commitEdit(doc, { type: 'cell', sheet: 'Summary', ref: 'B3', before: number(271219), after: number(2) })
    await firstSave
    assert.equal(session.editState().dirty, true)
    await doc.flush()
    assert.match(sheetText(host.bytes()), /<c r="B3"[^>]*><v>2<\/v>/)
  })
})

test('outside replacement discards a staged structural transaction and its snapshots', async () => {
  const book = new ExcelJS.Workbook()
  book.addWorksheet('Data').getCell('A1').value = 1
  const host = createHost(new Uint8Array(await book.xlsx.writeBuffer()))
  globalThis.orca = host.api
  try {
    const session = await import(`../src/ui/document.js?case=${Math.random()}`)
    const doc = await session.openDocument()
    const original = host.bytes()
    assert.equal(session.commitStructural(doc, { sheet: 'Data', axis: 'column', action: 'delete', index: 1, count: 1 }).ok, true)
    host.outsideWrite(original)
    assert.equal(session.editState().canUndo, false)
    assert.equal(session.undo(doc), null)
    await doc.flush()
    assert.deepEqual(host.bytes(), original)
  } finally {
    delete globalThis.orca
  }
})
