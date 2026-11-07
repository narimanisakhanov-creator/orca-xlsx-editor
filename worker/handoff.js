// The text a person hands to an agent: one context line naming the sheet and the range, then the
// table itself. Pure — no DOM, no Orca API — so `node --test` runs it directly, the editor frame
// imports it for the preview and the manual copy action, and the worker imports it to validate the
// one message it is allowed to send.
//
// The table body is not built here. The frame renders it with the same clipboard formatters a copy
// uses, so a handoff reads exactly like a paste; this module frames it, makes its fence safe for
// arbitrary cell values and bounds how much of a selection is materialized before the frame refuses
// to send and offers the full text for manual copying.

/**
 * `terminal.sendText` refuses longer text: Orca's host API schema is
 * `z.string().min(1).max(PANEL_ACTION_TEXT_MAX_LENGTH)`, and `PANEL_ACTION_TEXT_MAX_LENGTH` is 4096.
 * The handoff is deliberately one message: Orca offers no atomicity or idempotency across calls, so a
 * handoff longer than this is not split — it is not sent at all and the person copies it manually.
 */
export const TEXT_LIMIT = 4096

/** Terminal ids are opaque to us; the host caps them at 1024 characters. */
const TERMINAL_ID_MAX = 1024

// Ordered least-to-most settled among the states that still need attention, `done` last: an agent
// that is working is the likeliest target, one that finished is the least likely.
const STATE_RANK = { working: 0, waiting: 1, blocked: 2, done: 3 }
const UNKNOWN_STATE = 4

/**
 * A backtick run long enough not to be closed by any line of `body`: a cell whose value is ```` ``` ````
 * must not end the block and let the rest of the table escape it.
 */
function fenceMarker(body) {
  let longest = 0
  for (const line of body.split(/\r\n|\r|\n/)) {
    const match = /^[ \t]*(`{3,})/.exec(line)
    if (match) longest = Math.max(longest, match[1].length)
  }
  return '`'.repeat(Math.max(3, longest + 1))
}

/**
 * Handoff text for `terminal.sendText`.
 *
 * `cells` is the table body already rendered for the chosen `format` (Markdown from
 * `selectionToMarkdown`, TSV from `selectionToTsv`); an empty body means there is nothing to send
 * and comes back as an empty string, never as a header with no table. `format` only decides the
 * framing: a TSV body goes into a fenced block so an agent sees its tabs as one block instead of
 * loose columns; the fence is one backtick longer than any run at the start of a body line, so a
 * cell cannot close it early. A Markdown table needs no fence. An unknown format is treated as
 * Markdown, the default the popup offers.
 *
 * `fileName` is optional because a custom editor is handed bytes and learns no path; when a caller
 * does know the file it is named in the context line, and when it does not the line stays silent
 * rather than inventing one.
 */
export function buildHandoff({ sheetName, range, format, cells, fileName } = {}) {
  const body = typeof cells === 'string' ? cells : ''
  if (body === '') return ''
  // A line break in a name would put the rest of that name on its own line, where a backtick run
  // could open a code fence before the table does. The header stays one line.
  const flatten = (value) => (typeof value === 'string' ? value.replace(/[\r\n]+/g, ' ').trim() : '')
  const sheet = flatten(sheetName)
  const area = flatten(range)
  const file = flatten(fileName)

  const where = []
  if (sheet) where.push(`лист «${sheet}»`)
  if (area) where.push(`диапазон ${area}`)
  if (file) where.push(`файл «${file}»`)
  const header = where.length ? `Таблица из XLSX: ${where.join(', ')}` : 'Таблица из XLSX'

  if (format !== 'tsv') return `${header}\n\n${body}`
  const fence = fenceMarker(body)
  return `${header}\n\n${fence}tsv\n${body}\n${fence}`
}

/** Agents worth offering first: working, waiting, blocked, then done; newest state change first. */
export function rankAgents(agents) {
  const list = Array.isArray(agents) ? agents.filter((agent) => agent && typeof agent === 'object') : []
  const time = (agent) => (Number.isFinite(agent.stateStartedAt) ? agent.stateStartedAt : -Infinity)
  return list
    .map((agent, index) => ({ agent, index }))
    .sort((left, right) => {
      const leftRank = STATE_RANK[left.agent.state] ?? UNKNOWN_STATE
      const rightRank = STATE_RANK[right.agent.state] ?? UNKNOWN_STATE
      if (leftRank !== rightRank) return leftRank - rightRank
      const leftTime = time(left.agent)
      const rightTime = time(right.agent)
      if (leftTime !== rightTime) return rightTime - leftTime
      // Equal state and equal instant: keep the host's order so the list does not dance between
      // two refreshes that saw the same agents.
      return left.index - right.index
    })
    .map((entry) => entry.agent)
}

/**
 * How many leading rows fit the materialization budget. The frame counts row costs before building
 * the formatter matrix; when not every row fits, the full handoff is over the message limit, so the
 * frame refuses to send and copies the full selection only on demand.
 */
export function fitHandoffRows(rowCosts, budget) {
  const list = Array.isArray(rowCosts) ? rowCosts : []
  const cap = typeof budget === 'number' && budget >= 0 ? budget : Infinity
  let total = 0
  let rows = 0
  for (const cost of list) {
    const value = Number(cost)
    if (!Number.isFinite(value) || value < 0) return { rows, cut: true }
    if (total + value > cap) return { rows, cut: true }
    total += value
    rows++
  }
  return { rows, cut: false }
}

/**
 * Validates an `agent.send` request before it reaches the host, whose schema is strict and would
 * otherwise answer with a bare `invalid_params`. `enter` defaults to false: a handoff is text a
 * person reviews, and pressing Enter for them could submit half-written input to the agent.
 */
export function readSend(input) {
  if (!input || typeof input !== 'object') return { ok: false, reason: 'agent.send needs an object' }
  const terminalId = input.terminalId
  if (typeof terminalId !== 'string' || terminalId.length === 0) {
    return { ok: false, reason: 'Выберите агента с терминалом' }
  }
  if (terminalId.length > TERMINAL_ID_MAX) return { ok: false, reason: 'Идентификатор терминала слишком длинный' }
  const text = input.text
  if (typeof text !== 'string' || text.length === 0) return { ok: false, reason: 'Нечего передавать' }
  if (input.enter !== undefined && typeof input.enter !== 'boolean') {
    return { ok: false, reason: 'enter must be a boolean' }
  }
  return { ok: true, terminalId, text, enter: input.enter === true }
}