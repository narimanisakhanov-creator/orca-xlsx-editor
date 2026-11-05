import './ui/app.css'
import './ui/grid.css'

// Boot: wire the document session to the editor chrome and keep them in step.
import { openDocument, subscribeDocument, loadModel, setDateSystem } from './ui/document.js'
import { mountApp } from './ui/app.js'

const boot = document.getElementById('boot')

function status(text) {
  if (!boot) return
  boot.textContent = text
  boot.hidden = false
}

function hideBoot() {
  if (boot) boot.hidden = true
}

/**
 * Before the editor is mounted the boot screen is the only place to speak; after that, a message
 * belongs in the toolbar. A refusal must never cover the grid with a full-screen overlay.
 */
let app = null
function report(text, isError = false) {
  if (app) app.flash(text, isError)
  else status(text)
}

async function main() {
  status('Opening workbook…')
  const orca = globalThis.orca
  if (!orca?.document) {
    status('This editor needs Orca’s desktop app.')
    return
  }

  let rendering = false

  subscribeDocument({
    refused: (reason) => {
      report(`Could not save: ${reason}`, true)
    },
    accepted: () => {
      // The host took the edit: the tab is clean again.
      hideBoot()
      app?.onSaved?.()
    },
    discarded: () => {
      // The file changed outside while this tab had unsaved edits; the outside version wins, and
      // the undo history that described the old bytes is gone with it.
      app?.onDiscarded?.()
    }
  })

  const docApi = await openDocument()
  const opened = docApi.value
  if (opened?.zipStatus && opened.zipStatus !== 'ok') {
    // The archive is over the reader's budget (too many parts, too much inflated data, or broken).
    // Loading it anyway would unpack without a limit and can exhaust the renderer.
    status(`This workbook cannot be opened safely (${opened.zipStatus}).`)
    return
  }
  const first = await loadModel(opened.bytes)
  // Dates are stored as day numbers whose epoch depends on the file, so the writer needs to know it.
  setDateSystem(first.date1904)
  const root = document.getElementById('app')
  app = mountApp({ root, document: docApi, model: first, readOnly: docApi.readOnly })
  hideBoot()
  // editor/boot.js warns about missing assets if this flag is still unset after 8 s.
  window.__xlsxReady = true

  docApi.onReplace(async (next) => {
    // An outside write, or a reload after a conflict, replaces the document.
    if (rendering) return
    rendering = true
    try {
      status('Reloading from disk…')
      const model = await loadModel(next.bytes)
      setDateSystem(model.date1904)
      app.setModel(model)
      hideBoot()
    } catch (error) {
      // The bytes on disk are no longer a workbook this editor understands. Editing the stale model
      // would write its cells into that other file, so the tab goes read-only until it is reopened.
      app.block(`The file changed into something this editor cannot read: ${error instanceof Error ? error.message : String(error)}`)
    } finally {
      rendering = false
    }
  })
}

main().catch((error) => {
  status(`Could not open the workbook: ${error instanceof Error ? error.message : String(error)}`)
})
