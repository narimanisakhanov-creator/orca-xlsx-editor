// Search over displayed values and formulas: pure model behaviour, no DOM.
import { test } from 'node:test'
import assert from 'node:assert/strict'
import {
  searchSheet,
  searchWorkbook,
  continueSearch,
  nextSearchMatch,
  DEFAULT_SEARCH_LIMIT,
  MAX_SEARCH_QUERY
} from '../src/model/search.js'

const text = (display, formula = null) => ({ kind: 'text', value: display, display, formula })
const number = (value, display, formula = null) => ({ kind: 'number', value, display, formula })

/** A SheetModel-shaped fixture: `rows` is a grid of CellValue or null for an absent cell. */
function sheetOf(rows, { name = 'Sheet1', index = 0 } = {}) {
  const cells = new Map()
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      if (!value) return
      cells.set(`${r},${c}`, {
        ref: `${r},${c}`,
        row: r,
        col: c,
        value,
        css: { font: {}, fill: null, border: {}, align: {} }
      })
    })
  })
  return {
    name,
    index,
    rowCount: rows.length,
    colCount: rows.reduce((max, row) => Math.max(max, row.length), 0),
    cells,
    colWidths: new Map(),
    rowHeights: new Map(),
    merges: [],
    freeze: { rows: 0, cols: 0 },
    hidden: { rows: new Set(), cols: new Set() }
  }
}

function workbookOf(sheets) {
  return { sheets, active: 0, date1904: false, notes: [] }
}

test('searchWorkbook: display values across sheets in sheet, row, column order', () => {
  const first = sheetOf([[text('Alpha'), text('beta')], [text('ALPHA'), text('gamma')]], { name: 'One', index: 0 })
  const second = sheetOf([[text('alpha')]], { name: 'Two', index: 1 })
  const result = searchWorkbook(workbookOf([second, first]), 'alpha')

  assert.equal(result.reason, undefined)
  assert.equal(result.truncated, false)
  assert.equal(result.scannedCells, 5)
  assert.deepEqual(result.matches.map((match) => `${match.sheetIndex}:${match.sheetName}:${match.ref}:${match.field}`), [
    '0:One:A1:display',
    '0:One:A2:display',
    '1:Two:A1:display'
  ])
  assert.deepEqual(
    result.matches.map(({ offset, length }) => [offset, length]),
    [[0, 5], [0, 5], [0, 5]]
  )
})

test('searchSheet: formulas are searched with the leading = when the display does not match', () => {
  const sheet = sheetOf([[number(15, '15', 'SUM(B2:B6)'), text('plain')]])

  const byFragment = searchSheet(sheet, 'sum')
  assert.equal(byFragment.matches.length, 1)
  assert.deepEqual(byFragment.matches[0], {
    sheetIndex: sheet.index,
    sheetName: sheet.name,
    row: 0,
    col: 0,
    ref: 'A1',
    field: 'formula',
    text: '=SUM(B2:B6)',
    offset: 1,
    length: 3
  })

  const byEquals = searchSheet(sheet, '=SUM')
  assert.equal(byEquals.matches[0].field, 'formula')
  assert.deepEqual([byEquals.matches[0].offset, byEquals.matches[0].length], [0, 4])

  assert.equal(searchSheet(sheet, 'sum', { formulas: false }).matches.length, 0)
})

test('searchSheet: the displayed value wins when both the display and the formula match', () => {
  const value = { kind: 'text', value: 'SUM', display: 'SUM', formula: 'SUM(A1:A2)' }
  const result = searchSheet(sheetOf([[value]]), 'SUM')
  assert.equal(result.matches.length, 1)
  assert.equal(result.matches[0].field, 'display')
  assert.equal(result.matches[0].text, 'SUM')
})

test('search: repeated calls return equal, independent results and never mutate the model', () => {
  const sheet = sheetOf([[text('Alpha'), text('beta')], [text('alpha')]])
  const workbook = workbookOf([sheet])

  const first = searchWorkbook(workbook, 'alpha')
  const second = searchWorkbook(workbook, 'alpha')
  assert.deepEqual(first, second)

  first.matches[0].row = 99
  first.matches.pop()

  const third = searchWorkbook(workbook, 'alpha')
  assert.deepEqual(third, second)
  assert.equal(sheet.cells.get('0,0').value.display, 'Alpha')
})

test('search: Cyrillic and emoji match case-insensitively with offsets into the original text', () => {
  const sheet = sheetOf([[text('Магазин «Ко» & сын'), text('😀 Emoji')]])

  const shop = searchSheet(sheet, 'магазин')
  assert.equal(shop.matches[0].ref, 'A1')
  assert.deepEqual([shop.matches[0].offset, shop.matches[0].length], [0, 7])
  assert.equal(shop.matches[0].text, 'Магазин «Ко» & сын')

  const co = searchSheet(sheet, '«ко»')
  assert.deepEqual([co.matches[0].offset, co.matches[0].length], [8, 4])

  const amp = searchSheet(sheet, '&')
  assert.deepEqual([amp.matches[0].offset, amp.matches[0].length], [13, 1])

  const emoji = searchSheet(sheet, '😀')
  assert.equal(emoji.matches[0].ref, 'B1')
  assert.deepEqual([emoji.matches[0].offset, emoji.matches[0].length], [0, 2])
})

test('search: a lowercase expansion keeps offsets in the original string', () => {
  const sheet = sheetOf([[text('İstanbul')]])
  const result = searchSheet(sheet, 'stanbul')
  assert.equal(result.matches.length, 1)
  assert.deepEqual([result.matches[0].offset, result.matches[0].length], [1, 7])
})

test('search: a match inside one expanded character has a nonzero original span', () => {
  const inside = searchSheet(sheetOf([[text('İstanbul')]]), 'i')
  assert.equal(inside.matches.length, 1)
  assert.deepEqual([inside.matches[0].offset, inside.matches[0].length], [0, 1])

  // The whole-string lowercase keeps the context-sensitive final sigma; the range table still
  // maps the match back to the original characters.
  const greek = searchSheet(sheetOf([[text('İΟΣ')]]), 'ος')
  assert.equal(greek.matches.length, 1)
  assert.equal(greek.matches[0].text, 'İΟΣ')
  assert.deepEqual([greek.matches[0].offset, greek.matches[0].length], [1, 2])
})

test('search: caseSensitive and wholeCell narrow the same query', () => {
  const sheet = sheetOf([[text('Alpha'), text('alphabeta')]])

  assert.deepEqual(searchSheet(sheet, 'alpha').matches.map((match) => match.ref), ['A1', 'B1'])
  assert.deepEqual(searchSheet(sheet, 'alpha', { caseSensitive: true }).matches.map((match) => match.ref), ['B1'])
  assert.deepEqual(searchSheet(sheet, 'alpha', { wholeCell: true }).matches.map((match) => match.ref), ['A1'])
})

test('search: the match limit stops the walk and reports unexamined cells honestly', () => {
  const rows = Array.from({ length: 10 }, () => [text('hit'), text('hit'), text('miss')])
  const result = searchSheet(sheetOf(rows), 'hit', { limit: 3 })

  assert.equal(result.matches.length, 3)
  assert.equal(result.limitReached, true)
  // The walk stopped with 26 cells unexamined, so an earlier-coordinate match may still exist.
  assert.equal(result.truncated, true)
  assert.equal(result.scannedCells, 4)
  assert.equal(result.limit, 3)

  const lastMatch = searchSheet(sheetOf([[text('only')]]), 'only', { limit: 1 })
  assert.equal(lastMatch.matches.length, 1)
  assert.equal(lastMatch.limitReached, true)
  assert.equal(lastMatch.truncated, false)
  assert.equal(lastMatch.scannedCells, 1)
})

test('search: the scanned-cell ceiling bounds a query that finds nothing quickly', () => {
  const rows = Array.from({ length: 10 }, () => [text('hit'), text('hit'), text('miss')])
  const result = searchSheet(sheetOf(rows), 'miss', { maxCells: 2 })

  assert.equal(result.matches.length, 0)
  assert.equal(result.truncated, true)
  assert.equal(result.limitReached, false)
  assert.equal(result.scannedCells, 2)
})

test('search: truncated is false when the scan budget covered every cell', () => {
  const sheet = sheetOf([[text('x')], [text('y')]])
  const result = searchSheet(sheet, 'nope', { maxCells: 2 })

  assert.equal(result.matches.length, 0)
  assert.equal(result.scannedCells, 2)
  assert.equal(result.truncated, false)
  assert.equal(result.limitReached, false)
})

/** A Map that fails the test if a search reads more entries than its budget allows. */
class GuardedCells extends Map {
  constructor(entries, allowed) {
    super(entries)
    this.allowed = allowed
    this.read = 0
  }

  [Symbol.iterator]() {
    const inner = super[Symbol.iterator]()
    const self = this
    return {
      next() {
        const step = inner.next()
        if (!step.done) {
          self.read++
          if (self.read > self.allowed) throw new Error(`read past the ${self.allowed}-cell budget`)
        }
        return step
      }
    }
  }
}

test('search: maxCells bounds preparation, not just the matching loop', () => {
  const entries = Array.from({ length: 200 }, (_, row) => [
    `${row},0`,
    { ref: `${row},0`, row, col: 0, value: text('hit'), css: { font: {}, fill: null, border: {}, align: {} } }
  ])
  const cells = new GuardedCells(entries, 4)
  const sheet = { ...sheetOf([]), cells, rowCount: 200 }

  const result = searchSheet(sheet, 'hit', { maxCells: 3 })
  assert.equal(result.matches.length, 3)
  assert.equal(result.scannedCells, 3)
  assert.equal(result.truncated, true)
  assert.equal(cells.read, 4)
})

test('search: a cell edited into A1 after A2-A4 is reported as unexamined, never silently skipped', () => {
  // The map order after an edit is A2, A3, A4, then the appended A1 — not row-major order.
  const sheet = sheetOf([[null], [text('hit')], [text('hit')], [text('hit')]])
  sheet.cells.set('0,0', {
    ref: 'A1',
    row: 0,
    col: 0,
    value: text('hit'),
    css: { font: {}, fill: null, border: {}, align: {} }
  })

  const bounded = searchSheet(sheet, 'hit', { maxCells: 2 })
  assert.deepEqual(bounded.matches.map((match) => match.ref), ['A2', 'A3'])
  assert.equal(bounded.scannedCells, 2)
  assert.equal(bounded.truncated, true)
  assert.equal(bounded.limitReached, false)

  const full = searchSheet(sheet, 'hit', { maxCells: 10 })
  assert.deepEqual(full.matches.map((match) => match.ref), ['A1', 'A2', 'A3', 'A4'])
  assert.equal(full.truncated, false)
  assert.equal(full.limitReached, false)

  // Stopping at the first match found still leaves the appended A1 unexamined.
  const limited = searchSheet(sheet, 'hit', { maxCells: 10, limit: 1 })
  assert.deepEqual(limited.matches.map((match) => match.ref), ['A2'])
  assert.equal(limited.limitReached, true)
  assert.equal(limited.truncated, true)
})

test('search: a cursor continues into a later sheet without replaying the first one', () => {
  const first = sheetOf([[text('miss')]], { name: 'One', index: 0 })
  const second = sheetOf([[text('hit')]], { name: 'Two', index: 1 })
  const workbook = workbookOf([first, second])

  const page = searchWorkbook(workbook, 'hit', { maxCells: 1, revision: 'r1' })
  assert.deepEqual(page.matches, [])
  assert.equal(page.scannedCells, 1)
  assert.equal(page.truncated, true)
  assert.ok(page.cursor)

  const next = continueSearch(page.cursor, { revision: 'r1' })
  assert.deepEqual(next.matches.map((match) => `${match.sheetIndex}:${match.ref}`), ['1:A1'])
  assert.equal(next.scannedCells, 1)
  assert.equal(next.scannedTotal, 2)
  assert.equal(next.truncated, false)
  assert.equal(next.cursor, null)
})

test('search: a cursor moves past the match limit page by page', () => {
  const rows = Array.from({ length: 6 }, () => [text('hit')])
  const workbook = workbookOf([sheetOf(rows)])

  const page1 = searchWorkbook(workbook, 'hit', { limit: 2, maxCells: 100, revision: 'r1' })
  assert.deepEqual(page1.matches.map((match) => match.ref), ['A1', 'A2'])
  assert.equal(page1.limitReached, true)
  assert.equal(page1.truncated, true)
  assert.ok(page1.cursor)

  const page2 = continueSearch(page1.cursor, { revision: 'r1' })
  assert.deepEqual(page2.matches.map((match) => match.ref), ['A3', 'A4'])
  assert.equal(page2.scannedCells, 2)
  assert.equal(page2.scannedTotal, 4)
  assert.equal(page2.limitReached, true)
  assert.equal(page2.truncated, true)

  const page3 = continueSearch(page2.cursor, { revision: 'r1' })
  assert.deepEqual(page3.matches.map((match) => match.ref), ['A5', 'A6'])
  assert.equal(page3.scannedTotal, 6)
  assert.equal(page3.limitReached, true)
  assert.equal(page3.truncated, false)
  assert.equal(page3.cursor, null)
})

test('search: continuation reads only the remaining cells, never the prefix again', () => {
  const entries = Array.from({ length: 6 }, (_, row) => [
    `${row},0`,
    { ref: `${row},0`, row, col: 0, value: text('hit'), css: { font: {}, fill: null, border: {}, align: {} } }
  ])
  const cells = new GuardedCells(entries, 6)
  const sheet = { ...sheetOf([]), cells, rowCount: 6 }

  const page1 = searchSheet(sheet, 'hit', { maxCells: 2, revision: 'r1' })
  assert.equal(cells.read, 3)
  assert.deepEqual(page1.matches.map((match) => match.ref), ['A1', 'A2'])

  const page2 = continueSearch(page1.cursor, { revision: 'r1' })
  assert.equal(cells.read, 5)
  assert.deepEqual(page2.matches.map((match) => match.ref), ['A3', 'A4'])

  const page3 = continueSearch(page2.cursor, { revision: 'r1' })
  assert.equal(cells.read, 6)
  assert.deepEqual(page3.matches.map((match) => match.ref), ['A5', 'A6'])
  assert.equal(page3.cursor, null)
})

test('search: adding a cell invalidates the cursor instead of silently skipping it', () => {
  const sheet = sheetOf([[text('miss')], [text('miss')]])
  const workbook = workbookOf([sheet])

  const page = searchWorkbook(workbook, 'hit', { maxCells: 1, revision: 'r1' })
  assert.ok(page.cursor)

  // The edit lands at the end of the map, after the cursor, exactly like an app cell edit.
  sheet.cells.set('2,0', {
    ref: 'A3',
    row: 2,
    col: 0,
    value: text('hit'),
    css: { font: {}, fill: null, border: {}, align: {} }
  })

  const stale = continueSearch(page.cursor, { revision: 'r1' })
  assert.equal(stale.reason, 'cursor-stale')
  assert.equal(stale.cursor, null)
  assert.equal(stale.truncated, true)
  assert.deepEqual(stale.matches, [])

  const fresh = searchWorkbook(workbook, 'hit', { maxCells: 10 })
  assert.deepEqual(fresh.matches.map((match) => match.ref), ['A3'])
  assert.equal(fresh.truncated, false)
})

test('search: a revision token invalidates a cursor when the model generation changes', () => {
  const workbook = workbookOf([sheetOf([[text('miss')], [text('hit')]])])
  const page = searchWorkbook(workbook, 'hit', { maxCells: 1, revision: 'gen-1' })
  assert.ok(page.cursor)

  const stale = continueSearch(page.cursor, { revision: 'gen-2' })
  assert.equal(stale.reason, 'cursor-stale')
  assert.equal(stale.cursor, null)

  const other = searchWorkbook(workbook, 'hit', { maxCells: 1, revision: 'gen-3' })
  const same = continueSearch(other.cursor, { revision: 'gen-3' })
  assert.deepEqual(same.matches.map((match) => match.ref), ['A2'])
  assert.equal(same.cursor, null)
})

test('search: a truncated search without revision refuses to hand out a cursor', () => {
  const workbook = workbookOf([sheetOf([[text('miss')], [text('hit')]])])

  const page = searchWorkbook(workbook, 'hit', { maxCells: 1 })
  assert.equal(page.truncated, true)
  assert.equal(page.cursor, null)
  assert.equal(page.reason, 'revision-required')
  assert.deepEqual(page.matches, [])

  const withRevision = searchWorkbook(workbook, 'hit', { maxCells: 1, revision: 'r1' })
  assert.ok(withRevision.cursor)
  const next = continueSearch(withRevision.cursor, { revision: 'r1' })
  assert.deepEqual(next.matches.map((match) => match.ref), ['A2'])
})

test('search: continuation without revision is refused but does not consume the cursor', () => {
  const workbook = workbookOf([sheetOf([[text('miss')], [text('hit')]])])
  const page = searchWorkbook(workbook, 'hit', { maxCells: 1, revision: 'r1' })

  const missing = continueSearch(page.cursor)
  assert.equal(missing.reason, 'revision-required')
  assert.equal(missing.cursor, null)

  const retried = continueSearch(page.cursor, { revision: 'r1' })
  assert.deepEqual(retried.matches.map((match) => match.ref), ['A2'])
  assert.equal(retried.truncated, false)
})

test('search: edits to a pending or already visited cell make the cursor stale, never replay a value', () => {
  // Pending cell A2 held a match when it was probed; edit it in place (Map size stays the same).
  const sheet = sheetOf([[text('miss')], [text('hit')]])
  const pendingPage = searchSheet(sheet, 'hit', { maxCells: 1, revision: 1 })
  assert.ok(pendingPage.cursor)
  sheet.cells.get('1,0').value = text('miss')

  const pendingStale = continueSearch(pendingPage.cursor, { revision: 2 })
  assert.equal(pendingStale.reason, 'cursor-stale')
  assert.deepEqual(pendingStale.matches, [])
  assert.deepEqual(searchSheet(sheet, 'hit', { maxCells: 10, revision: 2 }).matches, [])

  // A match that was already returned by the previous page is edited in place.
  const visited = sheetOf([[text('hit')], [text('miss')]])
  const visitedPage = searchSheet(visited, 'hit', { maxCells: 1, revision: 1 })
  assert.deepEqual(visitedPage.matches.map((match) => match.ref), ['A1'])
  assert.ok(visitedPage.cursor)
  visited.cells.get('0,0').value = text('miss')

  const visitedStale = continueSearch(visitedPage.cursor, { revision: 2 })
  assert.equal(visitedStale.reason, 'cursor-stale')
  assert.deepEqual(visitedStale.matches, [])
  const visitedFresh = searchSheet(visited, 'hit', { maxCells: 10, revision: 2 })
  assert.deepEqual(visitedFresh.matches, [])
  assert.equal(visitedFresh.truncated, false)
})

test('search: a malformed cursor is reported, not guessed at', () => {
  assert.equal(continueSearch(null).reason, 'invalid-cursor')
  assert.equal(continueSearch({}).reason, 'invalid-cursor')
  assert.equal(continueSearch({ kind: 'search-cursor', walk: {}, sheets: [] }).reason, 'invalid-cursor')
})

test('search: empty and overlong queries report a reason instead of matching everything', () => {
  const sheet = sheetOf([[text('alpha')]])
  for (const query of ['', null, 42]) {
    const result = searchSheet(sheet, query)
    assert.equal(result.reason, 'empty-query')
    assert.deepEqual(result.matches, [])
  }
  const overlong = searchSheet(sheet, 'x'.repeat(MAX_SEARCH_QUERY + 1))
  assert.equal(overlong.reason, 'query-too-long')
  assert.deepEqual(overlong.matches, [])
})

test('search: sheetIndexes restricts a workbook search and a missing sheet is empty', () => {
  const first = sheetOf([[text('alpha')]], { name: 'One', index: 0 })
  const second = sheetOf([[text('alpha')]], { name: 'Two', index: 1 })
  const workbook = workbookOf([first, second])

  assert.deepEqual(searchWorkbook(workbook, 'alpha').matches.map((match) => match.sheetIndex), [0, 1])
  assert.deepEqual(
    searchWorkbook(workbook, 'alpha', { sheetIndexes: [1] }).matches.map((match) => match.sheetIndex),
    [1]
  )
  assert.deepEqual(searchWorkbook({ sheets: [] }, 'alpha').matches, [])
  assert.equal(searchSheet(null, 'alpha').scannedCells, 0)
})

test('nextSearchMatch: walks forward, wraps, and is stable when called twice', () => {
  const sheet = sheetOf([[text('x')], [text('x')], [text('x')]])
  const { matches } = searchSheet(sheet, 'x')
  assert.equal(matches.length, 3)
  assert.equal(DEFAULT_SEARCH_LIMIT, 1000)

  assert.equal(nextSearchMatch(matches, null), matches[0])
  assert.equal(nextSearchMatch(matches, matches[0]), matches[1])
  assert.equal(nextSearchMatch(matches, matches[2]), matches[0])
  assert.equal(nextSearchMatch(matches, matches[2], { wrap: false }), null)
  assert.equal(nextSearchMatch(matches, { sheetIndex: 0, row: 0, col: 0 }), matches[1])
  assert.equal(nextSearchMatch(matches, matches[0]), nextSearchMatch(matches, matches[0]))
  assert.equal(nextSearchMatch([], null), null)
  assert.equal(nextSearchMatch(null, null), null)
})