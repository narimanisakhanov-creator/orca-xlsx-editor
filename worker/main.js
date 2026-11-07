// The plugin's Node worker. It exists for one thing the editor frame cannot do itself:
// `terminal.sendText` is a worker-only host method, so the frame asks here and this module answers.
//
// It carries no formatting: the frame builds the handoff text (worker/handoff.js is shared, but the
// frame is what fills it in) and this module only validates, enforces the host's 4096-character
// limit and forwards exactly one call. Nothing is sent unless a person pressed the button.
//
// Orca offers no atomicity or idempotency between calls and no terminal draft state, so a handoff is
// never split into several messages: text over the limit is refused, not truncated or chunked.

import { TEXT_LIMIT, rankAgents, readSend } from './handoff.js'

/** A host failure is shown to the person verbatim; the code alone would not say what happened. */
function describe(error) {
  return error instanceof Error && error.message ? error.message : String(error)
}

/**
 * A throw with one of these codes happened while the host was validating or refusing the call, so it
 * provably typed nothing. Everything else — timeout in particular — may have reached the terminal
 * before failing; the person is warned that the result is unknown and nothing is retried.
 */
const NO_EFFECT_CODES = new Set([
  'invalid_request',
  'invalid_params',
  'unknown_method',
  'capability_denied',
  'consent_required',
  'not_permitted',
  'rate_limited',
  'too_large'
])

function hasKnownNoEffect(error) {
  const code = error?.code
  return typeof code === 'string' && NO_EFFECT_CODES.has(code)
}

export default function activate(orca) {
  // One handoff per terminal at a time. A second frame request while the first is still in flight is
  // refused: the first may still be typing, and a concurrent call could interleave its text.
  const sending = new Set()

  // The agents of the focused workspace, best target first. Sorted here rather than in the frame so
  // every caller of the bridge sees the same order.
  orca.panels.handle('agent.list', async () => {
    const result = await orca.host.call('agents.list')
    const agents = Array.isArray(result?.agents) ? result.agents : []
    return { agents: rankAgents(agents) }
  })

  // `{ terminalId, text, enter? }` in; `{ accepted, sent, unknown, reason? }` out. `sent` is exactly
  // the one message the host was given. `unknown` means the call failed in a way that does not prove
  // the terminal was untouched (a timeout may already have typed): the person must inspect it, and
  // the worker never retries.
  orca.panels.handle('agent.send', async (input) => {
    const request = readSend(input)
    if (!request.ok) {
      return { accepted: false, sent: '', unknown: false, reason: request.reason }
    }
    if (sending.has(request.terminalId)) {
      return { accepted: false, sent: '', unknown: false, reason: 'Передача в этот терминал уже идёт' }
    }
    sending.add(request.terminalId)
    try {
      if (request.text.length > TEXT_LIMIT) {
        return {
          accepted: false,
          sent: '',
          unknown: false,
          reason: `Текст длиннее ${TEXT_LIMIT} символов; отправка отменена — скопируйте полный текст вручную`
        }
      }

      try {
        const result = await orca.host.call('terminal.sendText', {
          terminalId: request.terminalId,
          text: request.text,
          enter: request.enter
        })
        if (result?.accepted === true) return { accepted: true, sent: request.text, unknown: false }
        if (result?.accepted === false) {
          return { accepted: false, sent: request.text, unknown: false, reason: 'Терминал не принял текст' }
        }
        return { accepted: false, sent: request.text, unknown: true, reason: 'Хост не подтвердил результат' }
      } catch (error) {
        // A timeout is the important case: Orca can report it after the text was typed, so the frame
        // must not imply the handoff failed. Validation/capability refusals are the only throws that
        // prove nothing was written.
        return {
          accepted: false,
          sent: request.text,
          unknown: !hasKnownNoEffect(error),
          reason: describe(error)
        }
      }
    } finally {
      sending.delete(request.terminalId)
    }
  })
}