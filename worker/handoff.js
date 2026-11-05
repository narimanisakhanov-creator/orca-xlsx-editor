// The text a person hands to an agent: one context line naming the sheet and the range, then the
// table itself. Pure — no DOM, no Orca API — so `node --test` runs it directly, the editor frame
// imports it for the preview and for what it sends, and the worker imports it to enforce the host's
// limit a second time.
//
// The table body is not built here. The frame renders it with the same clipboard formatters a copy
// uses, so a handoff reads exactly like a paste; this module only frames it and guards the size.

/**
 * `terminal.sendText` refuses longer text: Orca's host API schema is
 * `z.string().min(1).max(PANEL_ACTION_TEXT_MAX_LENGTH)`, and `PANEL_ACTION_TEXT_MAX_LENGTH` is 4096.
 */
export const TEXT_LIMIT = 4096

/** Terminal ids are opaque to us; the host caps them at 1024 characters. */
const TERMINAL_ID_MAX = 1024

// Ordered least-to-most settled among the states that still need attention, `done` last: an agent
// that is working is the likeliest target, one that finished is the least likely.
const STATE_RANK = { working: 0, waiting: 1, blocked: 2, done: 3 }
const UNKNOWN_STATE = 4

/**
 * Handoff text for `terminal.sendText`.
 *
 * `cells` is the table body already rendered for the chosen `format` (Markdown from
 * `selectionToMarkdown`, TSV from `selectionToTsv`); an empty body means there is nothing to send
 * and comes back as an empty string, never as a header with no table. `format` only decides the
 * framing: a TSV body goes into a fenced block so an agent sees its tabs as one block instead of
 * loose columns, a Markdown table needs no fence. An unknown format is treated as Markdown, the
 * default the popup offers.
 *
 * `fileName` is optional because a custom editor is handed bytes and learns no path; when a caller
 * does know the file it is named in the context line, and when it does not the line stays silent
 * rather than inventing one.
 */
export function buildHandoff({ sheetName, range, format, cells, fileName } = {}) {
  const body = typeof cells === 'string' ? cells : ''
  if (body === '') return ''
  const sheet = typeof sheetName === 'string' ? sheetName.trim() : ''
  const area = typeof range === 'string' ? range.trim() : ''
  const file = typeof fileName === 'string' ? fileName.trim() : ''

  const where = []
  if (sheet) where.push(`лист «${sheet}»`)
  if (area) where.push(`диапазон ${area}`)
  if (file) where.push(`файл «${file}»`)
  const header = where.length ? `Таблица из XLSX: ${where.join(', ')}` : 'Таблица из XLSX'

  if (format !== 'tsv') return `${header}\n\n${body}`
  return `${header}\n\n\`\`\`tsv\n${body}\n\`\`\``
}

/**
 * Keeps text inside the host's limit, cutting at a line boundary so a table is not sliced mid-row,
 * and says so in the text itself. The returned `text` is never longer than `limit`; `truncated`
 * tells the caller to report the loss instead of pretending the whole table arrived.
 */
export function truncateText(text, limit = TEXT_LIMIT) {
  const source = typeof text === 'string' ? text : ''
  if (!Number.isInteger(limit) || limit <= 0) return { text: '', truncated: source.length > 0 }
  if (source.length <= limit) return { text: source, truncated: false }

  const note = `\n…обрезано по лимиту ${limit} символов`
  const room = Math.max(0, limit - note.length)
  let cut = source.slice(0, room)
  const breakAt = cut.lastIndexOf('\n')
  // Only a break with content in front of it is worth cutting on; a single over-long first line is
  // cut hard, because an empty table would be worse than a broken one.
  if (breakAt > 0) cut = cut.slice(0, breakAt)
  return { text: (cut + note).slice(0, limit), truncated: true }
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
