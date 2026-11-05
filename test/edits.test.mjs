import assert from 'node:assert/strict'
import test from 'node:test'

import {
  createEditLog,
  parseRef,
  formatRef,
  shiftRef,
  columnName,
  equals
} from '../src/edit/edits.js'

/** @returns {{ kind: 'number', value: number, display: string, formula: null }} */
const num = (value, display = String(value)) => ({ kind: 'number', value, display, formula: null })
/** @returns {{ kind: 'text', value: string, display: string, formula: null }} */
const text = (value) => ({ kind: 'text', value, display: value, formula: null })

const cell = (sheet, ref, before, after) => ({ type: 'cell', sheet, ref, before, after })

/** Deterministic clock: each call advances 10 ms unless the test moves it by hand. */
function clock(start = 0, step = 10) {
  let t = start
  const now = () => t
  now.advance = (ms) => {
    t += ms
    return t
  }
  now.set = (value) => {
    t = value
    return t
  }
  now.step = step
  return now
}

test('apply, undo and redo keep order and content', () => {
  const log = createEditLog()
  const a = cell('Sheet1', 'A1', num(1), num(2))
  const b = cell('Sheet1', 'B2', text('x'), text('y'))
  const c = cell('Data', 'C3', null, num(10))

  assert.deepEqual(log.apply(a), a)
  assert.deepEqual(log.apply(b), b)
  assert.deepEqual(log.apply(c), c)

  assert.equal(log.ops().length, 3)
  assert.deepEqual(log.ops().map((op) => op.ref), ['A1', 'B2', 'C3'])
  assert.equal(log.canUndo, true)
  assert.equal(log.canRedo, false)

  assert.deepEqual(log.undo(), c)
  assert.deepEqual(log.undo(), b)
  assert.equal(log.canRedo, true)
  assert.deepEqual(log.ops().map((op) => op.ref), ['A1'])

  assert.deepEqual(log.undo(), a)
  assert.equal(log.undo(), null)
  assert.equal(log.canUndo, false)

  assert.deepEqual(log.redo(), a)
  assert.deepEqual(log.redo(), b)
  assert.deepEqual(log.redo(), c)
  assert.equal(log.redo(), null)
  assert.deepEqual(log.ops().map((op) => op.ref), ['A1', 'B2', 'C3'])
  assert.deepEqual(log.ops()[2].after, num(10))
})

test('a new edit clears the redo stack', () => {
  const log = createEditLog()
  log.apply(cell('Sheet1', 'A1', num(1), num(2)))
  log.apply(cell('Sheet1', 'A2', num(1), num(2)))
  log.undo()
  assert.equal(log.canRedo, true)

  const fresh = cell('Sheet1', 'A3', num(1), num(3))
  log.apply(fresh)

  assert.equal(log.canRedo, false)
  assert.equal(log.redo(), null)
  assert.deepEqual(log.ops().map((op) => op.ref), ['A1', 'A3'])
})

test('consecutive edits in one cell coalesce inside the window only', () => {
  const now = clock(0)
  const log = createEditLog({ now })

  log.apply(cell('Sheet1', 'B2', text(''), text('h')))
  now.advance(100)
  log.apply(cell('Sheet1', 'B2', text('h'), text('he')))
  now.advance(100)
  const merged = log.apply(cell('Sheet1', 'B2', text('he'), text('hel')))

  assert.equal(log.ops().length, 1)
  assert.deepEqual(merged, cell('Sheet1', 'B2', text(''), text('hel')))
  assert.deepEqual(log.ops()[0].before, text(''))
  assert.deepEqual(log.ops()[0].after, text('hel'))
  assert.equal(log.undoSize, 1)

  // Outside the window the two edits stay separate.
  now.advance(601)
  log.apply(cell('Sheet1', 'C2', text('a'), text('b')))
  now.advance(700)
  log.apply(cell('Sheet1', 'C2', text('b'), text('c')))
  assert.equal(log.ops().length, 3)
  assert.deepEqual(log.ops().slice(1).map((op) => op.after), [text('b'), text('c')])

  // A different cell in between also breaks the chain.
  const log2 = createEditLog({ now: clock(0) })
  log2.apply(cell('Sheet1', 'A1', num(1), num(2)))
  log2.apply(cell('Sheet1', 'A2', num(1), num(2)))
  log2.apply(cell('Sheet1', 'A1', num(2), num(3)))
  assert.equal(log2.ops().length, 3)

  // Coalescing can be switched off.
  const log3 = createEditLog({ now: clock(0), coalesce: false })
  log3.apply(cell('Sheet1', 'A1', num(1), num(2)))
  log3.apply(cell('Sheet1', 'A1', num(2), num(3)))
  assert.equal(log3.ops().length, 2)
})

test('an undo after a save is a new operation', () => {
  const saved = num(271219)
  const edited = num(1)
  const log = createEditLog({ now: clock(0) })

  log.apply(cell('Sheet1', 'B3', saved, edited))
  const mark = log.markSaved()

  assert.deepEqual(log.undo(), cell('Sheet1', 'B3', saved, edited))

  // The file holds 1 after the save, so the reversion has to be written, not dropped.
  const delta = log.opsSince(mark)
  assert.equal(delta.length, 1)
  assert.deepEqual(delta[0], cell('Sheet1', 'B3', edited, saved))
  assert.deepEqual(delta[0].after, saved)
  assert.deepEqual(delta[0].before, edited)

  // Once that delta is saved, nothing is pending and redo becomes a delta again.
  const mark2 = log.markSaved()
  assert.deepEqual(log.opsSince(mark2), [])

  assert.deepEqual(log.redo(), cell('Sheet1', 'B3', saved, edited))
  assert.deepEqual(log.opsSince(mark2), [cell('Sheet1', 'B3', saved, edited)])

  // Two undos across one save collapse into a single delta back to the original value.
  const log2 = createEditLog({ coalesce: false })
  log2.apply(cell('Sheet1', 'A1', num(1), num(2)))
  log2.apply(cell('Sheet1', 'A1', num(2), num(3)))
  const m2 = log2.markSaved()
  assert.deepEqual(log2.undo(), cell('Sheet1', 'A1', num(2), num(3)))
  assert.deepEqual(log2.undo(), cell('Sheet1', 'A1', num(1), num(2)))
  assert.deepEqual(log2.opsSince(m2), [cell('Sheet1', 'A1', num(3), num(1))])

  // Inside an unsaved batch the same undo is still a no-op for the writer.
  const log3 = createEditLog({ coalesce: false })
  const m3 = log3.markSaved()
  log3.apply(cell('Sheet1', 'D4', num(5), num(6)))
  log3.undo()
  assert.deepEqual(log3.opsSince(m3), [])
  assert.equal(log3.ops().length, 0)
})

test('coalescing still merges inside one batch and stops at a save', () => {
  const now = clock(0)
  const log = createEditLog({ now })

  log.apply(cell('Sheet1', 'A1', text('a'), text('ab')))
  now.advance(50)
  log.apply(cell('Sheet1', 'A1', text('ab'), text('abc')))
  assert.equal(log.ops().length, 1)

  const mark = log.markSaved()
  now.advance(50)
  log.apply(cell('Sheet1', 'A1', text('abc'), text('abcd')))

  // Well inside the 600 ms window, but the saved value is already in the file.
  assert.equal(log.ops().length, 2)
  assert.deepEqual(log.ops()[1], cell('Sheet1', 'A1', text('abc'), text('abcd')))
  assert.deepEqual(log.opsSince(mark), [cell('Sheet1', 'A1', text('abc'), text('abcd'))])

  // Undo closes the batch too: the tail record is not a merge target any more.
  const log2 = createEditLog({ now: clock(0) })
  log2.apply(cell('Sheet1', 'A1', num(1), num(2)))
  log2.apply(cell('Sheet1', 'A2', num(1), num(2)))
  log2.undo()
  log2.apply(cell('Sheet1', 'A1', num(2), num(9)))
  assert.equal(log2.ops().length, 2)
  assert.deepEqual(log2.ops()[0], cell('Sheet1', 'A1', num(1), num(2)))
  assert.deepEqual(log2.ops()[1], cell('Sheet1', 'A1', num(2), num(9)))
})

test('the undo stack keeps the last 1000 operations and reports the dropped ones', () => {
  const log = createEditLog({ coalesce: false })
  for (let i = 0; i < 1500; i++) {
    log.apply(cell('Sheet1', `A${i + 1}`, num(i), num(i + 1)))
  }

  assert.equal(log.ops().length, 1500)
  assert.equal(log.undoSize, 1000)
  assert.equal(log.dropped, 500)

  let undone = 0
  let op = log.undo()
  while (op !== null) {
    undone++
    if (undone === 1) assert.deepEqual(op.after, num(1500))
    op = log.undo()
  }
  assert.equal(undone, 1000)
  assert.equal(log.canUndo, false)
  assert.equal(log.undo(), null)
  assert.equal(log.dropped, 500)
  assert.equal(log.ops().length, 500)
  assert.deepEqual(log.ops()[0].ref, 'A1')
  assert.deepEqual(log.ops()[0].after, num(1))
  assert.deepEqual(log.ops()[499].ref, 'A500')
  assert.deepEqual(log.ops()[499].after, num(500))

  // The dropped tail is still redoable back into the window.
  let redone = 0
  while (log.redo() !== null) redone++
  assert.equal(redone, 1000)
  assert.equal(log.ops().length, 1500)
  assert.equal(log.undoSize, 1000)
})

test('opsSince returns only the delta since the last save mark', () => {
  const log = createEditLog({ coalesce: false })
  log.apply(cell('Sheet1', 'A1', num(1), num(2)))
  log.apply(cell('Sheet1', 'B1', text('a'), text('b')))

  const mark = log.markSaved()
  assert.deepEqual(log.opsSince(mark), [])

  log.apply(cell('Sheet1', 'C1', null, num(7)))
  log.apply(cell('Sheet1', 'D1', null, text('new')))
  const delta = log.opsSince(mark)
  assert.equal(delta.length, 2)
  assert.deepEqual(delta.map((op) => op.ref), ['C1', 'D1'])
  assert.deepEqual(delta[0], cell('Sheet1', 'C1', null, num(7)))

  // One cell edited many times after the mark shows up once, with the mark-time before.
  const mark2 = log.markSaved()
  log.apply(cell('Sheet1', 'E1', num(0), num(1)))
  log.apply(cell('Sheet1', 'E1', num(1), num(2)))
  log.apply(cell('Sheet1', 'E1', num(2), num(3)))
  const delta2 = log.opsSince(mark2)
  assert.equal(delta2.length, 1)
  assert.deepEqual(delta2[0], cell('Sheet1', 'E1', num(0), num(3)))

  // ... unless the cell ended up back at its mark-time value: then the delta is empty.
  const mark3 = log.markSaved()
  log.apply(cell('Sheet1', 'F1', num(5), num(6)))
  log.apply(cell('Sheet1', 'F1', num(6), num(5)))
  assert.deepEqual(log.opsSince(mark3), [])

  // Saving does not truncate the log: ops() still feeds a full rewrite.
  assert.equal(log.ops().length, 9)
})

test('a no-op edit is not recorded', () => {
  const log = createEditLog()
  const same = cell('Sheet1', 'A1', num(3), num(3))
  assert.equal(log.apply(same), null)
  assert.equal(log.ops().length, 0)
  assert.equal(log.canUndo, false)

  // Display text is derived, so it does not make an edit.
  assert.equal(log.apply(cell('Sheet1', 'A1', num(3, '3.00'), num(3, '3.0'))), null)
  assert.equal(log.ops().length, 0)

  // A changed formula does.
  const before = { kind: 'number', value: 4, display: '4', formula: 'A2*2' }
  const after = { kind: 'number', value: 4, display: '4', formula: 'A2*3' }
  assert.deepEqual(log.apply(cell('Sheet1', 'A1', before, after)), cell('Sheet1', 'A1', before, after))
  assert.equal(log.ops().length, 1)

  // Clearing an already blank cell is a no-op as well.
  const blank = { kind: 'blank', value: null, display: '', formula: null }
  assert.equal(log.apply(cell('Sheet1', 'B9', blank, blank)), null)
  assert.equal(log.ops().length, 1)
})

test('cell addresses round-trip, including the Z/AA and XFD edges', () => {
  assert.deepEqual(parseRef('B2'), { r: 1, c: 1 })
  assert.deepEqual(parseRef('$B$2'), { r: 1, c: 1 })
  assert.deepEqual(parseRef('a1'), { r: 0, c: 0 })
  assert.equal(formatRef({ r: 1, c: 1 }), 'B2')
  assert.equal(formatRef(parseRef('XFD1048576')), 'XFD1048576')
  assert.deepEqual(parseRef('XFD1048576'), { r: 1048575, c: 16383 })

  assert.equal(columnName(0), 'A')
  assert.equal(columnName(25), 'Z')
  assert.equal(columnName(26), 'AA')
  assert.equal(columnName(27), 'AB')
  assert.equal(columnName(51), 'AZ')
  assert.equal(columnName(52), 'BA')
  assert.equal(columnName(16383), 'XFD')

  assert.equal(formatRef({ r: 0, c: 16383 }), 'XFD1')
  assert.equal(shiftRef('Z1', 0, 1), 'AA1')
  assert.equal(shiftRef('AA1', 0, 1), 'AB1')
  assert.equal(shiftRef('AA10', 1, 0), 'AA11')
  assert.equal(shiftRef('B2', -1, -1), 'A1')
  assert.equal(shiftRef('XFD1048576', 0, 0), 'XFD1048576')

  assert.throws(() => shiftRef('A1', -1, 0), RangeError)
  assert.throws(() => shiftRef('XFD1', 0, 1), RangeError)
  assert.throws(() => parseRef('1A'), TypeError)
  assert.throws(() => parseRef(''), TypeError)
  assert.throws(() => parseRef('A0'), RangeError)
  assert.throws(() => columnName(16384), RangeError)
  assert.throws(() => columnName(-1), RangeError)

  assert.equal(equals(num(1), num(1)), true)
  assert.equal(equals(num(1), num(1, '1.00')), true)
  assert.equal(equals(num(1), num(2)), false)
  assert.equal(equals(text('a'), num(1)), false)
  assert.equal(equals(null, null), true)
  assert.equal(equals(null, num(0)), false)
  const err = { kind: 'error', value: { code: '#DIV/0!' }, display: '#DIV/0!', formula: '1/0' }
  const errCopy = { kind: 'error', value: { code: '#DIV/0!' }, display: '#DIV/0!', formula: '1/0' }
  assert.equal(equals(err, errCopy), true)
})

test('undo after coalescing rewinds to the value before the first keystroke', () => {
  const now = clock(0)
  const log = createEditLog({ now })

  log.apply(cell('Sheet1', 'A1', text('keep'), text('x')))
  now.advance(50)
  log.apply(cell('Sheet1', 'A1', text('x'), text('xy')))
  now.advance(50)
  log.apply(cell('Sheet1', 'A1', text('xy'), text('xyz')))

  assert.equal(log.ops().length, 1)

  const undone = log.undo()
  assert.deepEqual(undone.before, text('keep'))
  assert.deepEqual(undone.after, text('xyz'))
  assert.deepEqual(log.opsSince(log.markSaved()), [])

  const redone = log.redo()
  assert.deepEqual(redone.after, text('xyz'))
  assert.equal(log.ops().length, 1)

  // Typing back to the pre-edit value inside the window removes the edit entirely.
  now.advance(50)
  log.undo()
  log.apply(cell('Sheet1', 'A1', text('keep'), text('q')))
  now.advance(50)
  // Typing back to the pre-edit value settles the cell on `keep`. The log drops the operation, but
  // it still reports that settled value: the grid must not keep showing the keystroke that is gone.
  const collapsed = log.apply(cell('Sheet1', 'A1', text('q'), text('keep')))
  assert.equal(collapsed.collapsed, true)
  assert.deepEqual(collapsed.after, text('keep'))
  assert.equal(log.ops().length, 0)
  assert.equal(log.canUndo, false)
})

test('clear resets the log and keeps old save marks meaningful', () => {
  const log = createEditLog({ coalesce: false })
  log.apply(cell('Sheet1', 'A1', num(1), num(2)))
  const mark = log.markSaved()
  log.apply(cell('Sheet1', 'A2', num(1), num(2)))

  log.clear()
  assert.equal(log.ops().length, 0)
  assert.equal(log.undoSize, 0)
  assert.equal(log.dropped, 0)
  assert.equal(log.canUndo, false)
  assert.equal(log.canRedo, false)
  assert.deepEqual(log.opsSince(mark), [])

  log.apply(cell('Sheet1', 'A3', num(1), num(2)))
  assert.equal(log.opsSince(mark).length, 1)
  assert.equal(log.undo().ref, 'A3')
})

test('invalid operations and payloads are rejected', () => {
  const log = createEditLog()
  assert.throws(() => log.apply(null), TypeError)
  assert.throws(() => log.apply({ type: 'style', sheet: 'S', ref: 'A1' }), TypeError)
  assert.throws(() => log.apply({ type: 'cell', sheet: '', ref: 'A1' }), TypeError)
  assert.throws(() => log.apply({ type: 'cell', sheet: 'S', ref: 'nope' }), TypeError)
  assert.throws(() => log.apply(cell('S', 'A1', { kind: 'bogus', value: 1 }, num(2))), TypeError)
  assert.throws(() => log.opsSince(undefined), TypeError)
  assert.throws(() => createEditLog({ undoLimit: 0 }), RangeError)
  assert.throws(() => createEditLog({ coalesceWindowMs: -1 }), RangeError)
})
