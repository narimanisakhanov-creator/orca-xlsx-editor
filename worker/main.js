// The plugin's Node worker. It exists for one thing the editor frame cannot do itself:
// `terminal.sendText` is a worker-only host method, so the frame asks here and this module answers.
//
// It carries no formatting: the frame builds the handoff text (worker/handoff.js is shared, but the
// frame is what fills it in) and this module only validates, enforces the host's text limit and
// forwards the call. Nothing is sent unless a person pressed the button in the frame.

import { TEXT_LIMIT, rankAgents, readSend, truncateText } from './handoff.js'

/** A host failure is shown to the person verbatim; the code alone would not say what happened. */
function describe(error) {
  return error instanceof Error && error.message ? error.message : String(error)
}

export default function activate(orca) {
  // The agents of the focused workspace, best target first. Sorted here rather than in the frame so
  // every caller of the bridge sees the same order.
  orca.panels.handle('agent.list', async () => {
    const result = await orca.host.call('agents.list')
    const agents = Array.isArray(result?.agents) ? result.agents : []
    return { agents: rankAgents(agents) }
  })

  // `{ terminalId, text, enter? }` in; `{ accepted, sent, truncated, reason? }` out. `sent` is what
  // the terminal was actually given, so a frame can tell a full handoff from a cut one.
  orca.panels.handle('agent.send', async (input) => {
    const request = readSend(input)
    if (!request.ok) return { accepted: false, sent: '', truncated: false, reason: request.reason }

    const { text, truncated } = truncateText(request.text, TEXT_LIMIT)
    if (text.length === 0) return { accepted: false, sent: '', truncated, reason: 'Нечего передавать' }

    try {
      const result = await orca.host.call('terminal.sendText', {
        terminalId: request.terminalId,
        text,
        enter: request.enter
      })
      if (result?.accepted === true) return { accepted: true, sent: text, truncated }
      return { accepted: false, sent: text, truncated, reason: 'Терминал не принял текст' }
    } catch (error) {
      // A terminal that closed between the listing and the click rejects here; the person gets the
      // host's reason instead of a request that never resolves.
      return { accepted: false, sent: text, truncated, reason: describe(error) }
    }
  })
}
