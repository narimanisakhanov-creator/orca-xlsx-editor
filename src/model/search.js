// Search across a workbook's displayed values and formulas. Pure: no DOM, no Orca API, so
// `node --test` runs it directly and a caller can re-run it on every keystroke without touching
// the document. Matches are plain records sorted by sheet, row and column, so "find next" is a
// pure walk over the previous result instead of state kept somewhere.
//
// Search is bounded on purpose, and its bound is honest. The sheet map is insertion-ordered, and
// an edit appends the changed cell even when its coordinates belong earlier, so map order is not
// proof of row-major order. A bounded walk therefore makes no "first by coordinates" claim: it
// reads a prefix, sorts what it found, and sets `truncated` whenever at least one materialised
// cell was not examined. Only `truncated === false` means the returned matches are exhaustive.
// A caller that needs globally ordered find-next after edits should obtain a maintained
// coordinate-ordered cell index from the model/edits layer instead of raising the budget.
//
// The budget is per page, not per search: a result carries an opaque cursor holding the live sheet
// iterator position, and `continueSearch` resumes exactly there, so a 200,000-cell sheet or a
// 10,000-match page does not make the rest of the workbook unreachable. A cursor is refused after
// any cell Map changes identity or size, or when the caller's `revision` differs; restart then.

/** How many matches a search returns before it stops and reports `truncated`. */
export const DEFAULT_SEARCH_LIMIT = 1000
/** Upper bound a caller's `limit` is clamped to, so one query cannot return an unbounded list. */
export const MAX_SEARCH_LIMIT = 10000
/** Longest query worth folding; Excel's own Find box is 255 characters. */
export const MAX_SEARCH_QUERY = 256
/** Materialised cells a single search walks before it stops and reports `truncated`. */
export const MAX_SEARCH_CELLS = 200000

/**
 * @typedef {'display'|'formula'} SearchField
 * @typedef {{ sheetIndex: number, sheetName: string, row: number, col: number, ref: string,
 *   field: SearchField, text: string, offset: number, length: number }} SearchMatch
 * @typedef {object} SearchCursor Opaque continuation token; pass the one from the latest result
 *   to `continueSearch`, never edit or reuse it after the model changed.
 * @typedef {{ query: string, matches: SearchMatch[], truncated: boolean, limitReached: boolean,
 *   scannedCells: number, scannedTotal: number, limit: number, cursor: SearchCursor | null,
 *   reason?: 'empty-query'|'query-too-long'|'revision-required'|'cursor-stale'|'invalid-cursor' }} SearchResult
 *
 * `truncated` means at least one materialised cell was not examined: the matches are a sorted
 * subset and cannot be reported as "everything" or as "the earliest cells". This can be caused by
 * the scan budget or by stopping at `limit`. `limitReached` separately says the walk stopped
 * because exactly `limit` matches came back. When `truncated` is false the walk covered every
 * materialised cell, so `matches` is exhaustive and coordinate-sorted.
 *
 * A non-null `cursor` is the bounded continuation point: it holds the sheet and the live Map
 * iterator position, so `continueSearch` resumes after the last examined cell instead of walking
 * the prefix again. Continuation requires a `revision` token on the search call and the same token
 * on every `continueSearch`: the owner must advance it on edit, paste, undo, redo and reload.
 * Without a token a truncated search returns `reason: 'revision-required'` and no cursor, because
 * a value edit to an existing cell does not change the Map size and cannot be detected otherwise.
 * Even with a token, a sheet's cell Map being replaced or changing size makes the cursor stale.
 */

/**
 * Search one sheet. Same options and result shape as `searchWorkbook`.
 *
 * @param {object} sheet
 * @param {string} query
 * @param {{ limit?: number, maxCells?: number, caseSensitive?: boolean, wholeCell?: boolean,
 *   formulas?: boolean, revision?: unknown }} [options]
 * @returns {SearchResult}
 */
export function searchSheet(sheet, query, options = {}) {
  return startSearch([sheet], query, options)
}

/**
 * Search every sheet of a workbook, or the sheets named by `options.sheetIndexes`.
 *
 * A cell matches when its displayed text matches, or, unless `options.formulas === false`, when
 * its formula matches. Formula text is matched with Excel's leading `=`, so `=SUM` finds the
 * formula and `SUM` finds it too (at offset 1). One cell contributes at most one match; the
 * displayed value wins when both sides contain the query.
 *
 * UI contract: `matches` is always sorted by `sheetIndex`, `row`, `col`, and `nextSearchMatch`
 * walks that list. If `truncated` is true, the list is a bounded sample — never show "no more
 * matches", "nothing found", or navigation that implies earlier cells were considered. When
 * `cursor` is non-null, `continueSearch(cursor)` resumes after the last examined cell; offer that
 * as "search further" instead of asking the user to widen `maxCells`. If `limitReached` is true,
 * the same continuation also moves past the match limit. Cursors from different pages are not
 * globally sorted; sort the concatenation when the UI needs sheet order. Continuation requires a
 * `revision` token here and the same token on every `continueSearch`; a truncated search without
 * one returns `reason: 'revision-required'` and a null cursor.
 *
 * @param {object} workbook
 * @param {string} query
 * @param {{ limit?: number, maxCells?: number, caseSensitive?: boolean, wholeCell?: boolean,
 *   formulas?: boolean, sheetIndexes?: number[], revision?: unknown }} [options]
 * @returns {SearchResult}
 */
export function searchWorkbook(workbook, query, options = {}) {
  const sheets = Array.isArray(workbook?.sheets) ? workbook.sheets : []
  const wanted = Array.isArray(options.sheetIndexes) ? new Set(options.sheetIndexes) : null
  const selected = sheets
    .filter((sheet) => sheet && (!wanted || wanted.has(sheet.index)))
    .sort((a, b) => Number(a.index ?? 0) - Number(b.index ?? 0))
  return startSearch(selected, query, options)
}

/**
 * Continue a bounded search from the latest cursor. The query and matching options come from the
 * cursor; only the work bounds may change. Returns another `SearchResult` whose `cursor` is the
 * next continuation point, or `null` when the walk reached the end.
 *
 * `revision` is mandatory and must equal the token supplied when the search started. A missing
 * token is reported as `revision-required` without consuming the cursor, so the caller can retry
 * with the right generation. A different token makes the cursor unusable (`cursor-stale`), as
 * does a replaced or resized cell Map. Restart with `searchWorkbook`/`searchSheet` then; the new
 * pass includes the edited cells.
 *
 * @param {SearchCursor} cursor
 * @param {{ limit?: number, maxCells?: number, revision: unknown }} options
 * @returns {SearchResult}
 */
export function continueSearch(cursor, options = {}) {
  options = options ?? {}
  if (!isCursor(cursor) || cursor.invalid || cursor.exhausted) return rejectedResult(cursor, 'invalid-cursor')
  if (cursor.revision === undefined) {
    cursor.invalid = true
    return rejectedResult(cursor, 'revision-required')
  }
  if (options.revision === undefined) return rejectedResult(cursor, 'revision-required')
  if (options.revision !== cursor.revision || cursorStale(cursor)) {
    cursor.invalid = true
    return rejectedResult(cursor, 'cursor-stale')
  }
  const limit = normalizeLimit(options.limit ?? cursor.limit)
  const maxCells = normalizeMaxCells(options.maxCells ?? cursor.maxCells)
  return advance(cursor, limit, maxCells)
}

/**
 * The match after `from`, in result order, wrapping to the first one unless `wrap` is false.
 * `from` is a match or any `{ sheetIndex, row, col }` position; `null` returns the first match.
 * Pure: calling it twice with the same arguments returns the same match.
 *
 * @param {SearchMatch[]} matches
 * @param {{ sheetIndex?: number, row: number, col: number } | SearchMatch | null} [from]
 * @param {{ wrap?: boolean }} [options]
 * @returns {SearchMatch | null}
 */
export function nextSearchMatch(matches, from = null, { wrap = true } = {}) {
  const list = Array.isArray(matches) ? matches.filter(isMatch) : []
  if (list.length === 0) return null
  if (!from || !Number.isInteger(from.row) || !Number.isInteger(from.col)) return list[0]

  const position = {
    sheetIndex: Number.isInteger(from.sheetIndex) ? from.sheetIndex : 0,
    row: from.row,
    col: from.col
  }
  const next = list.find((match) => compareMatches(match, position) > 0)
  if (next) return next
  return wrap ? list[0] : null
}

function startSearch(sheets, query, options) {
  const limit = normalizeLimit(options.limit)
  const maxCells = normalizeMaxCells(options.maxCells)
  const prepared = prepareQuery(query, options)
  if (prepared.reason) return reasonResult(query, prepared.reason, limit)

  const cursor = {
    kind: 'search-cursor',
    query,
    prepared,
    sheets,
    // Positions are kept in insertion order: a bounded walk is allowed to stop anywhere without
    // first proving that map order matches sheet order.
    walk: { sheets, sheetIndex: 0, iterator: null, pending: null },
    snapshot: sheets.map(snapshotOf),
    limit,
    maxCells,
    revision: options.revision,
    scannedTotal: 0,
    invalid: false,
    exhausted: false
  }
  return advance(cursor, cursor.limit, cursor.maxCells)
}

/**
 * One bounded page. The cursor keeps the live Map iterator, so the next call continues after the
 * last examined entry instead of replaying the prefix. `peekRaw` tells `truncated` whether any
 * materialised cell is still unexamined, including a cell on a later sheet.
 */
function advance(cursor, limit, maxCells) {
  cursor.limit = limit
  cursor.maxCells = maxCells
  const result = {
    query: cursor.query,
    matches: [],
    truncated: false,
    limitReached: false,
    scannedCells: 0,
    scannedTotal: cursor.scannedTotal,
    limit,
    cursor: null
  }

  let hitLimit = false
  let hasMore = false
  for (;;) {
    if (result.scannedCells >= maxCells) {
      hasMore = peekRaw(cursor.walk)
      break
    }
    const raw = takeRaw(cursor.walk)
    if (!raw) break

    const entry = entryOf(raw.key, raw.view)
    if (!entry) continue
    const [row, col, view] = entry
    result.scannedCells++
    cursor.scannedTotal++

    const match = matchCell(raw.sheet, view, row, col, cursor.prepared)
    if (!match) continue
    result.matches.push(match)
    if (result.matches.length >= limit) {
      hitLimit = true
      hasMore = peekRaw(cursor.walk)
      break
    }
  }

  result.scannedTotal = cursor.scannedTotal
  result.limitReached = hitLimit
  result.truncated = hasMore
  if (!hasMore) {
    cursor.exhausted = true
  } else if (cursor.revision === undefined) {
    // A value edit to an existing cell does not change the Map size. Without a generation token
    // the cursor could silently return a stale value, so refuse continuation instead.
    cursor.invalid = true
    result.reason = 'revision-required'
  } else {
    result.cursor = cursor
  }
  // Insertion order is not sheet order, so a page only promises coordinate order for itself.
  result.matches.sort(compareMatches)
  return result
}

/** The next raw cell in insertion order, crossing sheets, or `null` at the end of the walk. */
function takeRaw(walk) {
  if (walk.pending) {
    const raw = walk.pending
    walk.pending = null
    // Pending is a key, not a snapshot: an in-place edit to that cell must be visible instead of
    // replaying the value that was probed before the edit.
    return { sheet: raw.sheet, key: raw.key, view: raw.sheet?.cells?.get?.(raw.key) }
  }
  while (walk.sheetIndex < walk.sheets.length) {
    const sheet = walk.sheets[walk.sheetIndex]
    const cells = sheet?.cells
    if (!cells) {
      walk.sheetIndex++
      continue
    }
    if (!walk.iterator) walk.iterator = cells[Symbol.iterator]()
    const step = walk.iterator.next()
    if (!step.done) return { sheet, key: step.value[0], view: step.value[1] }
    walk.iterator = null
    walk.sheetIndex++
  }
  return null
}

/** Whether one more materialised cell exists, leaving it pending so continuation keeps it. */
function peekRaw(walk) {
  if (walk.pending) return true
  const raw = takeRaw(walk)
  if (!raw) return false
  walk.pending = { sheet: raw.sheet, key: raw.key }
  return true
}

function snapshotOf(sheet) {
  const cells = sheet?.cells ?? null
  return {
    sheet,
    cells,
    size: typeof cells?.size === 'number' ? cells.size : null
  }
}

/** A cursor dies when a cell Map is replaced or grows/shrinks; `revision` covers in-place edits. */
function cursorStale(cursor) {
  for (const snapshot of cursor.snapshot) {
    const cells = snapshot.sheet?.cells ?? null
    if (cells !== snapshot.cells) return true
    const size = typeof cells?.size === 'number' ? cells.size : null
    if (size !== snapshot.size) return true
  }
  return false
}

function isCursor(cursor) {
  return Boolean(
    cursor &&
    cursor.kind === 'search-cursor' &&
    cursor.walk &&
    Array.isArray(cursor.walk.sheets) &&
    Number.isInteger(cursor.walk.sheetIndex) &&
    Number.isInteger(cursor.scannedTotal) &&
    cursor.prepared &&
    Array.isArray(cursor.sheets) &&
    Array.isArray(cursor.snapshot)
  )
}

/** No walk happened at all (empty or overlong query), so nothing is unexamined. */
function reasonResult(query, reason, limit = DEFAULT_SEARCH_LIMIT) {
  return {
    query: typeof query === 'string' ? query : '',
    matches: [],
    truncated: false,
    limitReached: false,
    scannedCells: 0,
    scannedTotal: 0,
    limit,
    cursor: null,
    reason
  }
}

/** The cursor cannot continue: it is stale or malformed, so the result is incomplete by definition. */
function rejectedResult(cursor, reason) {
  return {
    query: typeof cursor?.query === 'string' ? cursor.query : '',
    matches: [],
    truncated: true,
    limitReached: false,
    scannedCells: 0,
    scannedTotal: cursor?.scannedTotal ?? 0,
    limit: cursor?.limit ?? DEFAULT_SEARCH_LIMIT,
    cursor: null,
    reason
  }
}

function entryOf(key, view) {
  const fallback = String(key).split(',')
  const row = Number(view?.row ?? fallback[0])
  const col = Number(view?.col ?? fallback[1])
  if (!Number.isInteger(row) || !Number.isInteger(col)) return null
  return [row, col, view]
}

function prepareQuery(query, options) {
  if (typeof query !== 'string' || query === '') return { reason: 'empty-query' }
  if (query.length > MAX_SEARCH_QUERY) return { reason: 'query-too-long' }
  const caseSensitive = options.caseSensitive === true
  return {
    query,
    needle: caseSensitive ? query : query.toLowerCase(),
    caseSensitive,
    wholeCell: options.wholeCell === true,
    formulas: options.formulas !== false
  }
}

function matchCell(sheet, view, row, col, prepared) {
  const value = view?.value
  if (!value) return null

  const display = value.display == null ? '' : String(value.display)
  const displayHit = findMatch(display, prepared)
  if (displayHit) {
    return makeMatch(sheet, row, col, 'display', display, displayHit)
  }

  if (prepared.formulas && value.formula) {
    const formula = String(value.formula)
    const text = formula.startsWith('=') ? formula : `=${formula}`
    const formulaHit = findMatch(text, prepared)
    if (formulaHit) return makeMatch(sheet, row, col, 'formula', text, formulaHit)
  }
  return null
}

/**
 * First occurrence of the prepared query, as `{ offset, length }` in the original text.
 * Folding can change a string's length (`İ` lowercases to two code units), so the mapping back to
 * original offsets is built only when the lengths differ.
 */
function findMatch(text, prepared) {
  if (text === '') return null
  if (prepared.caseSensitive) {
    if (prepared.wholeCell) return text === prepared.query ? { offset: 0, length: text.length } : null
    const offset = text.indexOf(prepared.query)
    return offset === -1 ? null : { offset, length: prepared.query.length }
  }

  const folded = foldText(text)
  if (prepared.wholeCell) {
    return folded.text === prepared.needle ? { offset: 0, length: text.length } : null
  }
  const index = folded.text.indexOf(prepared.needle)
  if (index === -1) return null
  if (!folded.starts) return { offset: index, length: prepared.needle.length }
  // A match that ends inside a folded expansion (İ -> i + combining dot) still belongs to the
  // whole original character, so both boundaries come from per-character ranges.
  const start = folded.starts[index]
  const end = folded.ends[index + prepared.needle.length - 1]
  return { offset: start, length: end - start }
}

/**
 * Lowercased text to search, plus per-folded-unit original ranges when a character expands.
 * Lowercasing the whole string keeps context-sensitive mappings (final Greek sigma) correct;
 * the range maps are built per code point because only they know which original character an
 * expanded unit came from.
 */
function foldText(text) {
  const lower = text.toLowerCase()
  if (lower.length === text.length) return { text: lower, starts: null, ends: null }

  const starts = []
  const ends = []
  let original = 0
  for (const char of text) {
    const part = char.toLowerCase()
    for (let index = 0; index < part.length; index++) {
      starts.push(original)
      ends.push(original + char.length)
    }
    original += char.length
  }
  // A whole-string mapping that disagrees with the per-character lengths is outside what the
  // range table can describe; direct offsets stay correct for the text as it lowercased.
  if (starts.length !== lower.length) return { text: lower, starts: null, ends: null }
  return { text: lower, starts, ends }
}

function makeMatch(sheet, row, col, field, text, hit) {
  return {
    sheetIndex: Number(sheet.index ?? 0),
    sheetName: String(sheet.name ?? ''),
    row,
    col,
    ref: `${columnName(col)}${row + 1}`,
    field,
    text,
    offset: hit.offset,
    length: hit.length
  }
}

function columnName(col) {
  let name = ''
  let value = col + 1
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

function isMatch(value) {
  return value && Number.isInteger(value.row) && Number.isInteger(value.col)
}

function compareMatches(a, b) {
  return (Number(a.sheetIndex ?? 0) - Number(b.sheetIndex ?? 0))
    || (a.row - b.row)
    || (a.col - b.col)
}

function normalizeLimit(value) {
  const number = Number(value)
  if (!Number.isFinite(number) || number <= 0) return DEFAULT_SEARCH_LIMIT
  return Math.min(Math.floor(number), MAX_SEARCH_LIMIT)
}

function normalizeMaxCells(value) {
  const number = Number(value)
  if (!Number.isFinite(number) || number <= 0) return MAX_SEARCH_CELLS
  return Math.min(Math.floor(number), MAX_SEARCH_CELLS)
}