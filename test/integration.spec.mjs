// End-to-end check of the document layer without Orca: a fake host that keeps the file like
// Orca's session does, so a real edit can be saved and re-read from bytes.
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import path from 'node:path'
import { readZip } from '../src/write/zip.js'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = path.join(HERE, 'fixtures')

/**
 * Stands in for `orca.document.session` with the same contract: one edit in flight, debounce,
 * a version per accepted write, and an outside write that can arrive first.
 */
function fakeHost(bytes, options = {}) {
  const state = {
    version: 1,
    bytes: new Uint8Array(bytes),
    readOnly: Boolean(options.readOnly),
    edits: [],
    discards: 0
  }
  const listeners = new Set()
  const handlers = { decode: null, encode: null }
  let timer = null
  let inFlight = null
  let dirty = false

  function deliver() {
    const content = { kind: 'bytes', bytes: state.bytes.buffer.slice(0) }
    return { version: state.version, readOnly: state.readOnly, content }
  }

  function applyEdit(edit) {
    return new Promise((resolve, reject) => {
      queueMicrotask(() => {
        state.edits.push(edit)
        if (edit.baseVersion !== state.version) {
          const error = new Error('conflict')
          error.code = 'conflict'
          reject(error)
          return
        }
        state.bytes = new Uint8Array(edit.content.bytes)
        state.version += 1
        resolve()
      })
    })
  }

  function send() {
    clearTimeout(timer)
    const encoded = handlers.encode()
    return applyEdit({ baseVersion: state.version, content: encoded }).then(
      () => {
        dirty = false
        handlers.decode(encoded)
      },
      (error) => {
        state.discards += 1
        throw error
      }
    )
  }

  const api = {
    document: {
      get: async () => deliver(),
      onDidChange() {
        return { dispose() {} }
      },
      applyEdit,
      session: async (options2) => {
        handlers.decode = options2.decode
        handlers.encode = options2.encode
        handlers.decode(deliver().content)
        return {
          get value() {
            return handlers.decode(deliver().content)
          },
          get readOnly() {
            return state.readOnly
          },
          change() {
            if (state.readOnly) return
            dirty = true
            clearTimeout(timer)
            timer = setTimeout(() => {
              if (!inFlight) inFlight = send().finally(() => (inFlight = null))
            }, 0)
          },
          async flush() {
            clearTimeout(timer)
            const current = inFlight || (dirty ? send() : null)
            if (current) await current
          },
          onReplace(fn) {
            listeners.add(fn)
            return { dispose: () => listeners.delete(fn) }
          }
        }
      }
    }
  }

  return { api, state, listeners, outsideWrite(next) { state.bytes = new Uint8Array(next) } }
}

async function withHost(file, fn) {
  const bytes = readFileSync(path.join(FIXTURES, file))
  const host = fakeHost(bytes)
  globalThis.orca = host.api
  try {
    const mod = await import(`../src/ui/document.js?case=${Math.random()}`)
    return await fn(mod, host, bytes)
  } finally {
    delete globalThis.orca
  }
}

/** The parts of a workbook as name to bytes, through the plugin's own reader (no bare imports, so
 *  `orca plugin test` can run this file too). */
function parts(bytes) {
  const result = readZip(new Uint8Array(bytes))
  if (result.status !== 'ok') throw new Error(`unreadable zip: ${result.status}`)
  return Object.fromEntries(result.entries.map((entry) => [entry.name, entry.data]))
}

test('an edit saves through the session and only the edited sheet changes', async () => {
  await withHost('macro.xlsm', async (doc, host, original) => {
    const session = await doc.openDocument()
    const before = Object.fromEntries(Object.entries(parts(original)).map(([k, v]) => [k, Buffer.from(v)]))

    const result = doc.commitEdit(session, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: { kind: 'number', value: 282349 },
      after: { kind: 'number', value: 999999 }
    })
    assert.equal(result.ok, true, result.reason)
    await session.flush()
    await new Promise((resolve) => setTimeout(resolve, 20))

    const after = parts(host.state.bytes)

    // The macro project and the styles must survive byte for byte.
    assert.deepEqual(Buffer.from(after['xl/vbaProject.bin']), before['xl/vbaProject.bin'])
    assert.deepEqual(Buffer.from(after['xl/styles.xml']), before['xl/styles.xml'])
    if (before['xl/sharedStrings.xml']) {
      assert.deepEqual(Buffer.from(after['xl/sharedStrings.xml']), before['xl/sharedStrings.xml'])
    }

    // Part names are unchanged: nothing added, nothing dropped.
    assert.deepEqual(Object.keys(after).sort(), Object.keys(before).sort())

    const sheet = Buffer.from(after['xl/worksheets/sheet1.xml']).toString('utf8')
    assert.match(sheet, /<v>999999<\/v>/)
  })
})

test('a text edit round-trips with XML-special characters', async () => {
  await withHost('rich.xlsx', async (doc, host) => {
    const session = await doc.openDocument()
    const text = 'A & B < C > D "E" \'F\''
    doc.commitEdit(session, {
      type: 'cell',
      sheet: 'Notes',
      ref: 'A2',
      before: { kind: 'blank', value: null },
      after: { kind: 'text', value: text }
    })
    await session.flush()
    await new Promise((resolve) => setTimeout(resolve, 20))

    const sheet = Buffer.from(parts(host.state.bytes)['xl/worksheets/sheet2.xml']).toString('utf8')
    assert.match(sheet, /A &amp; B &lt; C &gt; D/)
    assert.doesNotMatch(sheet, /<t xml:space="preserve">A & B/)
  })
})

test('undo rewinds without leaving a change in the file', async () => {
  await withHost('rich.xlsx', async (doc, host) => {
    const session = await doc.openDocument()
    const before = Buffer.from(host.state.bytes)

    doc.commitEdit(session, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B3',
      before: { kind: 'number', value: 271219 },
      after: { kind: 'number', value: 1 }
    })
    await session.flush()
    await new Promise((resolve) => setTimeout(resolve, 20))

    const changed = Buffer.from(host.state.bytes)
    assert.notDeepEqual(changed, before)

    doc.undo(session)
    await session.flush()
    await new Promise((resolve) => setTimeout(resolve, 20))

    const restored = parts(host.state.bytes)
    const sheet = Buffer.from(restored['xl/worksheets/sheet1.xml']).toString('utf8')
    assert.match(sheet, /<v>271219<\/v>/)
    assert.doesNotMatch(sheet, /<v>1<\/v>/)
  })
})

test('state reports dirty until the host accepts', async () => {
  await withHost('rich.xlsx', async (doc, host) => {
    const session = await doc.openDocument()
    assert.equal(doc.editState().dirty, false)
    doc.commitEdit(session, {
      type: 'cell',
      sheet: 'Summary',
      ref: 'B2',
      before: { kind: 'number', value: 282349 },
      after: { kind: 'number', value: 5 }
    })
    assert.equal(doc.editState().dirty, true)
  })
})
