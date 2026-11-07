// The editor chrome around the grid: sheet tabs, formula bar, status bar, the copy/export menus
// and the keyboard. All Orca calls go through ./document.js; nothing here touches bytes directly.

import { createGrid } from './grid.js'
import { commitEdit, commitStructural, currentDocumentBytes, undo, redo, editState, loadModel, setDateSystem } from './document.js'
import { selectionToTsv, selectionToCsv, selectionToMarkdown, selectionToHtml, selectionFitsTextBudget, copyText, copyHtml } from './clipboard.js'
import { searchWorkbook, continueSearch } from '../model/search.js'
import { buildHandoff, fitHandoffRows, TEXT_LIMIT } from '../../worker/handoff.js'

const ROWS_AHEAD = 2

// A handoff materializes a selection the way an export does, so both use one budget: past it the
// frame refuses before building a huge intermediate matrix.
const EXPORT_CELL_LIMIT = 200_000
// Bound the complete manual handoff before creating its formatter matrix or clipboard copies.
const HANDOFF_COPY_CHAR_LIMIT = 8 * 1024 * 1024

export function mountApp({ root, document: docApi, model, readOnly }) {
  const el = (tag, className, text) => {
    const node = window.document.createElement(tag)
    if (className) node.className = className
    if (text != null) node.textContent = text
    return node
  }

  root.textContent = ''

  const shell = el('div', 'wb')
  const toolbar = el('div', 'wb-toolbar')
  const tabs = el('div', 'wb-tabs')
  const spacer = el('div', 'wb-toolbar-spacer')
  const undoButton = el('button', 'wb-button', 'Undo')
  const redoButton = el('button', 'wb-button', 'Redo')
  const stateLabel = el('span', 'wb-state')

  const menus = []

  function menuItem(label, shortcut, action) {
    const item = el('button', 'wb-menu-item', label)
    item.type = 'button'
    item.setAttribute('role', 'menuitem')
    if (shortcut) {
      const hint = el('span', 'wb-menu-shortcut', shortcut)
      hint.setAttribute('aria-hidden', 'true')
      item.append(hint)
    }
    item.addEventListener('click', () => {
      closeMenus()
      action()
      // The menu is closed and its items are hidden: focus belongs back on the grid.
      grid.focus?.()
    })
    return item
  }

  // A plain popup: one open menu at a time, closed by Escape or a pointer press outside it.
  function makeMenu({ label, ariaLabel }) {
    const wrap = el('div', 'wb-menu')
    const button = el('button', 'wb-button wb-menu-button', label)
    button.type = 'button'
    button.setAttribute('aria-label', ariaLabel)
    button.setAttribute('aria-haspopup', 'menu')
    button.setAttribute('aria-expanded', 'false')
    const popup = el('div', 'wb-menu-popup')
    popup.setAttribute('role', 'menu')
    popup.hidden = true
    wrap.append(button, popup)

    const menu = {
      wrap,
      button,
      isOpen: () => !popup.hidden,
      close() {
        if (popup.hidden) return
        popup.hidden = true
        button.setAttribute('aria-expanded', 'false')
      },
      open() {
        for (const other of menus) if (other !== menu) other.close()
        popup.hidden = false
        button.setAttribute('aria-expanded', 'true')
        popup.querySelector('button, input')?.focus()
      },
      toggle() {
        if (menu.isOpen()) menu.close()
        else menu.open()
      },
      add(item) {
        popup.append(item)
        return item
      }
    }
    button.addEventListener('click', () => menu.toggle())
    menus.push(menu)
    return menu
  }

  function closeMenus() {
    for (const menu of menus) menu.close()
  }

  let copyFormulas = false
  const copyMenu = makeMenu({ label: 'Копировать', ariaLabel: 'Копировать выделение' })
  copyMenu.add(menuItem('Копировать', 'Ctrl+C', () => copySelection(selectionToTsv)))
  copyMenu.add(menuItem('С форматированием', '', copyFormattedSelection))
  copyMenu.add(menuItem('Как Markdown', '', () => copySelection(selectionToMarkdown)))
  copyMenu.add(menuItem('Как CSV', '', () => copySelection(selectionToCsv)))
  const formulaToggle = el('label', 'wb-menu-check')
  const formulaBox = el('input')
  formulaBox.type = 'checkbox'
  formulaBox.setAttribute('aria-label', 'Формулы как текст')
  formulaToggle.append(formulaBox, el('span', null, 'Формулы как текст'))
  formulaBox.addEventListener('change', () => {
    copyFormulas = formulaBox.checked
    // A handoff is built from the same text a copy would produce, so the preview has to follow.
    if (agentMenu?.isOpen()) renderHandoffPreview()
  })
  copyMenu.add(formulaToggle)

  const exportMenu = makeMenu({ label: 'Экспорт', ariaLabel: 'Экспортировать лист в файл' })
  exportMenu.add(menuItem('CSV (.csv)', '', () => exportSheet('csv')))
  exportMenu.add(menuItem('Markdown (.md)', '', () => exportSheet('markdown')))

  // Structural edits are transactional: the writer preflights the whole package and refuses
  // unsupported references before anything changes, and the history can undo them even after a save.
  const structureMenu = makeMenu({ label: 'Строки', ariaLabel: 'Вставить или удалить строки и столбцы' })
  structureMenu.add(menuItem('Вставить строки выше', '', () => structuralChange('insert', 'row')))
  structureMenu.add(menuItem('Удалить строки', '', () => structuralChange('delete', 'row')))
  structureMenu.add(menuItem('Вставить столбцы слева', '', () => structuralChange('insert', 'column')))
  structureMenu.add(menuItem('Удалить столбцы', '', () => structuralChange('delete', 'column')))

  // "В агента": `terminal.sendText` is a worker-only host method, so the button exists only when
  // this frame can reach its own worker. Nothing is ever sent on its own — the host requires a user
  // gesture, and the click on "Передать" is that gesture.
  const workerBridge = globalThis.orca?.plugin
  const canHandoff = Boolean(workerBridge && typeof workerBridge.request === 'function')

  const HANDOFF_FORMATS = [['markdown', 'Markdown-таблица'], ['tsv', 'TSV']]
  const AGENT_STATE_TEXT = { working: 'работает', waiting: 'ожидает', blocked: 'нужен ответ', done: 'готово' }
  const AGENT_STATES = new Set(Object.keys(AGENT_STATE_TEXT))

  let handoffFormat = 'markdown'
  let handoffAgents = []
  let handoffAgentId = null
  let handoffLoading = false
  let handoffNote = ''
  let handoffList = null
  let handoffEmpty = null
  let handoffPreview = null
  let handoffNotice = null
  let handoffCopy = null
  let handoffSend = null
  let handoffEnter = null
  let handoffFormatSelect = null
  // The built handoff for the current selection and format; `{ key, plan }` or null after any edit.
  let handoffCache = null
  // One send is awaiting the worker; preview re-renders must not re-enable the button. Orca has no
  // idempotency, so a second concurrent call to the same terminal must not start.
  let handoffSending = false

  const agentMenu = canHandoff ? makeMenu({ label: 'В агента', ariaLabel: 'Передать таблицу агенту' }) : null

  if (agentMenu) {
    const popup = agentMenu.wrap.querySelector('.wb-menu-popup')
    // The popup holds a list, a select, a checkbox and a button: a form, not a menu.
    popup.classList.add('wb-handoff-popup')
    popup.setAttribute('role', 'dialog')
    popup.setAttribute('aria-label', 'Передать таблицу агенту')

    handoffList = el('div', 'wb-agent-list')
    handoffList.setAttribute('role', 'radiogroup')
    handoffList.setAttribute('aria-label', 'Агенты этого workspace')

    handoffEmpty = el('div', 'wb-empty', 'Нет активных агентов в этом workspace')
    handoffEmpty.hidden = true

    const formatRow = el('label', 'wb-menu-check')
    const formatSelect = el('select', 'wb-select')
    handoffFormatSelect = formatSelect
    formatSelect.setAttribute('aria-label', 'Формат таблицы')
    for (const [value, text] of HANDOFF_FORMATS) {
      const option = el('option', null, text)
      option.value = value
      formatSelect.append(option)
    }
    formatSelect.value = handoffFormat
    formatSelect.addEventListener('change', () => {
      handoffFormat = HANDOFF_FORMATS.some(([value]) => value === formatSelect.value) ? formatSelect.value : 'markdown'
      renderHandoffPreview()
    })
    formatRow.append(el('span', null, 'Формат'), formatSelect)

    const enterRow = el('label', 'wb-menu-check')
    handoffEnter = el('input')
    handoffEnter.type = 'checkbox'
    handoffEnter.setAttribute('aria-label', 'Отправить сразу (Enter)')
    enterRow.append(handoffEnter, el('span', null, 'Отправить сразу (Enter)'))

    handoffPreview = el('pre', 'wb-handoff-preview')
    handoffPreview.setAttribute('aria-label', 'Предпросмотр того, что уйдёт агенту')

    // Explains why a handoff is not being sent; the preview itself stays exactly what would be sent.
    handoffNotice = el('div', 'wb-empty')
    handoffNotice.setAttribute('role', 'status')
    handoffNotice.hidden = true

    // The full table is copied through the same legacy copy command the Copy menu uses. It is the
    // manual path when the text does not fit the single-message limit.
    handoffCopy = el('button', 'wb-button wb-handoff-copy', 'Скопировать полный текст')
    handoffCopy.type = 'button'
    handoffCopy.setAttribute('aria-label', 'Скопировать полный текст таблицы для ручной вставки')
    handoffCopy.addEventListener('click', copyFullHandoff)

    handoffSend = el('button', 'wb-button wb-handoff-send', 'Передать')
    handoffSend.type = 'button'
    handoffSend.disabled = true
    handoffSend.addEventListener('click', () => {
      void sendHandoff()
    })

    agentMenu.add(handoffList)
    agentMenu.add(handoffEmpty)
    agentMenu.add(formatRow)
    agentMenu.add(enterRow)
    agentMenu.add(handoffPreview)
    agentMenu.add(handoffNotice)
    agentMenu.add(handoffCopy)
    agentMenu.add(handoffSend)

    // The agent list is read when the popup opens: an agent may have started, finished or gone
    // since the last look, and the frame has no subscription to keep it current.
    agentMenu.button.addEventListener('click', () => {
      if (!agentMenu.isOpen()) return
      // A file may have reloaded or an edit landed since the last preview; rebuild from the model.
      handoffCache = null
      renderHandoffPreview()
      void refreshHandoffAgents()
    })
  }

  function agentName(agent) {
    const title = typeof agent?.title === 'string' ? agent.title.trim() : ''
    if (title) return title
    const kind = typeof agent?.agent === 'string' ? agent.agent.trim() : ''
    if (kind) return kind
    return typeof agent?.terminalId === 'string' && agent.terminalId !== '' ? agent.terminalId : 'Без названия'
  }

  /** Draws the rows the worker returned: already ranked there, so the order is kept as it came. */
  function renderHandoffAgents() {
    if (!handoffList) return
    handoffList.textContent = ''
    for (const agent of handoffAgents) {
      const id = typeof agent?.terminalId === 'string' ? agent.terminalId : ''
      const row = el('button', 'wb-agent')
      row.type = 'button'
      row.setAttribute('role', 'radio')
      const checked = id !== '' && id === handoffAgentId
      row.setAttribute('aria-checked', String(checked))
      if (checked) row.classList.add('is-selected')
      const dot = el('span', 'wb-agent-dot')
      dot.dataset.state = AGENT_STATES.has(agent?.state) ? agent.state : 'unknown'
      row.append(
        dot,
        el('span', 'wb-agent-name', agentName(agent)),
        el('span', 'wb-agent-state', AGENT_STATE_TEXT[agent?.state] ?? '')
      )
      if (!id) {
        // Orca could not tie this agent's pane to a terminal it lists, so there is no address to
        // type into; offering the row would only produce a refusal on click.
        row.disabled = true
        row.title = 'У этого агента нет терминала, которому можно передать текст'
      } else {
        row.addEventListener('click', () => {
          handoffAgentId = id
          renderHandoffAgents()
          renderHandoffPreview()
        })
      }
      handoffList.append(row)
    }
    if (handoffAgents.length === 0) {
      handoffEmpty.hidden = false
      handoffEmpty.textContent = handoffNote || 'Нет активных агентов в этом workspace'
    } else {
      handoffEmpty.hidden = true
    }
  }

  function pickHandoffAgent() {
    const usable = handoffAgents.filter((agent) => typeof agent?.terminalId === 'string' && agent.terminalId !== '')
    if (usable.some((agent) => agent.terminalId === handoffAgentId)) return
    handoffAgentId = usable[0]?.terminalId ?? null
  }

  async function refreshHandoffAgents() {
    if (!agentMenu || handoffLoading) return
    handoffLoading = true
    // A list already on screen stays there while it refreshes; only an empty one says why it is.
    handoffNote = handoffAgents.length === 0 ? 'Ищу агентов…' : ''
    renderHandoffAgents()
    try {
      const result = await workerBridge.request('agent.list')
      handoffAgents = Array.isArray(result?.agents) ? result.agents : []
      handoffNote = ''
    } catch (error) {
      handoffAgents = []
      handoffNote = 'Не удалось получить список агентов'
      flash(describeError(error), true)
    } finally {
      handoffLoading = false
      pickHandoffAgent()
      renderHandoffAgents()
      renderHandoffPreview()
      // Rebuilding the rows drops the node that had focus; without this the popup would be open
      // with focus on the page body and Escape or Tab would go somewhere else.
      if (agentMenu.isOpen() && !agentMenu.wrap.contains(window.document.activeElement)) {
        const target = handoffList.querySelector('.wb-agent:not(:disabled)') ?? handoffFormatSelect ?? handoffSend
        target?.focus?.()
      }
    }
  }

  /** The box the handoff covers: the selection, or the active cell when nothing is selected. */
  function handoffBox() {
    const sel = selectionBox()
    if (sel) {
      return {
        r1: Math.min(sel.r1, sel.r2),
        c1: Math.min(sel.c1, sel.c2),
        r2: Math.max(sel.r1, sel.r2),
        c2: Math.max(sel.c1, sel.c2)
      }
    }
    const parsed = parseRef(active.ref)
    return parsed ? { r1: parsed.r, c1: parsed.c, r2: parsed.r, c2: parsed.c } : null
  }

  function handoffRange(box) {
    const start = formatRef({ r: box.r1, c: box.c1 })
    const end = formatRef({ r: box.r2, c: box.c2 })
    return start === end ? start : `${start}:${end}`
  }

  /** The inputs the plan depends on; any edit clears the cache explicitly instead of in the key. */
  function handoffKey(box) {
    return `${active.sheet}\u0000${box.r1},${box.c1},${box.r2},${box.c2}\u0000${handoffFormat}\u0000${copyFormulas}`
  }

  /**
   * A lower bound on formatted characters, stopping once `TEXT_LIMIT` is exceeded. It must never
   * overestimate: otherwise a message that fits could be refused before formatting.
   */
  function handoffRowCosts(sheet, box) {
    const costs = []
    let total = 0
    for (let r = box.r1; r <= box.r2; r++) {
      let cost = 0
      for (let c = box.c1; c <= box.c2; c++) {
        const value = sheet.cells.get(`${r},${c}`)?.value
        const text = copyFormulas && value?.formula ? `=${value.formula}` : (value?.display ?? '')
        let length = text.length
        if (handoffFormat === 'tsv' && !(copyFormulas && value?.formula)) {
          for (const match of text.matchAll(/[\t\r\n]+/g)) length -= match[0].length - 1
        }
        cost += length + (handoffFormat === 'tsv' ? 1 : 3)
      }
      // TSV has one fewer separator than cells across the whole matrix.
      if (handoffFormat === 'tsv' && r === box.r1) cost--
      costs.push(cost)
      total += cost
      if (total > TEXT_LIMIT) break
    }
    return { costs, complete: costs.length === box.r2 - box.r1 + 1 }
  }

  /**
   * The one message the agent could receive: the table formatted once, never split. `text` is
   * exactly what would be sent and the preview shows it; `oversized` means the selection does not
   * fit one message, so nothing is sent and the person copies the full text manually. A row count
   * that already exceeds the character budget is not formatted at all.
   */
  function handoffPlan() {
    const sheet = model.sheets[active.sheet]
    const box = handoffBox()
    if (!sheet || !box) {
      return { key: '', text: '', oversized: false, error: 'Нечего передавать' }
    }
    const size = (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1)
    if (size > EXPORT_CELL_LIMIT) {
      return { key: '', text: '', oversized: false, error: `Выделение слишком велико (${size} ячеек; максимум ${EXPORT_CELL_LIMIT}). Уменьшите выделение.` }
    }
    const key = handoffKey(box)
    if (handoffCache?.key === key) return handoffCache.plan

    const { costs, complete } = handoffRowCosts(sheet, box)
    const fit = fitHandoffRows(costs, TEXT_LIMIT)
    let plan
    if (!complete || fit.cut) {
      plan = { key, text: '', oversized: true, error: '' }
    } else {
      const cells = handoffFormat === 'tsv'
        ? selectionToTsv(sheet, box, { formulas: copyFormulas })
        : selectionToMarkdown(sheet, box, { formulas: copyFormulas })
      const text = buildHandoff({
        sheetName: sheet.name,
        range: handoffRange(box),
        format: handoffFormat,
        cells
      })
      if (text === '') plan = { key, text: '', oversized: false, error: 'Нечего передавать' }
      else if (text.length > TEXT_LIMIT) plan = { key, text, oversized: true, error: '' }
      else plan = { key, text, oversized: false, error: '' }
    }
    handoffCache = { key, plan }
    return plan
  }

  function renderHandoffPreview() {
    if (!handoffPreview) return
    const plan = handoffPlan()
    handoffPreview.textContent = plan.error || plan.oversized ? '' : plan.text
    let notice = ''
    if (plan.error) {
      notice = plan.error
    } else if (plan.oversized) {
      notice = `Полный текст больше ${TEXT_LIMIT} символов — отправка отключена. Скопируйте полный текст и вставьте вручную.`
    }
    handoffNotice.textContent = notice
    handoffNotice.hidden = notice === ''
    handoffSend.disabled =
      Boolean(plan.error) || plan.oversized || handoffSending ||
      plan.text === '' || handoffAgentId === null
  }

  /** The full, untruncated handoff for a manual paste; same formatters and fence as the preview. */
  function copyFullHandoff() {
    const sheet = model.sheets[active.sheet]
    const box = handoffBox()
    if (!sheet || !box) {
      flash('Нечего копировать', true)
      return
    }
    const size = (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1)
    if (size > EXPORT_CELL_LIMIT) {
      flash(`Выделение слишком велико (${size} ячеек; максимум ${EXPORT_CELL_LIMIT}). Уменьшите выделение.`, true)
      return
    }
    const headerAllowance = String(sheet.name ?? '').length + handoffRange(box).length + 128
    if (!selectionFitsTextBudget(sheet, box, {
      formulas: copyFormulas,
      maxChars: HANDOFF_COPY_CHAR_LIMIT - headerAllowance
    })) {
      flash('Полный текст слишком велик для копирования. Уменьшите выделение; ничего не скопировано.', true)
      return
    }
    const cells = handoffFormat === 'tsv'
      ? selectionToTsv(sheet, box, { formulas: copyFormulas })
      : selectionToMarkdown(sheet, box, { formulas: copyFormulas })
    const text = buildHandoff({
      sheetName: sheet.name,
      range: handoffRange(box),
      format: handoffFormat,
      cells
    })
    if (text === '') {
      flash('Нечего копировать', true)
      return
    }
    const copied = copyText(text)
    flash(copied ? 'Полный текст скопирован' : 'Не удалось скопировать', !copied)
  }

  async function sendHandoff() {
    if (!agentMenu || handoffSending) return
    const agent = handoffAgents.find((entry) => entry?.terminalId === handoffAgentId) ?? null
    const plan = handoffPlan()
    if (!agent) {
      flash('Выберите агента', true)
      return
    }
    if (plan.error) {
      flash(plan.error, true)
      return
    }
    if (plan.oversized) {
      flash(`Текст больше ${TEXT_LIMIT} символов — отправка отключена. Скопируйте полный текст вручную.`, true)
      return
    }
    if (plan.text === '') {
      flash('Нечего передавать', true)
      return
    }
    if (handoffPreview.textContent !== plan.text) {
      renderHandoffPreview()
      flash('Предпросмотр изменился. Проверьте его перед отправкой.', true)
      return
    }
    handoffSending = true
    handoffSend.disabled = true
    try {
      const result = await workerBridge.request('agent.send', {
        terminalId: agent.terminalId,
        text: plan.text,
        enter: handoffEnter.checked === true
      })
      if (result?.accepted === true) {
        flash(`Передано в ${agentName(agent)}`)
        agentMenu.close()
        grid.focus?.()
      } else if (result?.unknown === true) {
        flash(
          `Результат неизвестен: ${result.reason || 'хост не подтвердил отправку'}. Проверьте терминал — автоматического повтора нет.`,
          true
        )
      } else {
        flash(result?.reason || 'Агент не принял текст', true)
      }
    } catch (error) {
      // The bridge failed too; the call may still have reached the terminal, so do not claim failure.
      flash(
        `Результат неизвестен: ${describeError(error)}. Проверьте терминал — автоматического повтора нет.`,
        true
      )
    } finally {
      handoffSending = false
      // The popup stays open so the person can check the terminal or retry deliberately.
      renderHandoffAgents()
      renderHandoffPreview()
    }
  }

  const onOutsidePointer = (event) => {
    for (const menu of menus) {
      if (menu.isOpen() && !menu.wrap.contains(event.target)) menu.close()
    }
  }
  const onMenuEscape = (event) => {
    if (event.key !== 'Escape') return
    for (const menu of menus) {
      if (!menu.isOpen()) continue
      event.preventDefault()
      menu.close()
      menu.button.focus()
    }
  }
  window.document.addEventListener('pointerdown', onOutsidePointer)
  window.document.addEventListener('keydown', onMenuEscape)

  const searchButton = el('button', 'wb-button', 'Найти')
  searchButton.type = 'button'
  searchButton.setAttribute('aria-label', 'Поиск по книге (Ctrl+F)')
  searchButton.setAttribute('aria-expanded', 'false')
  searchButton.setAttribute('aria-controls', 'wb-search-panel')

  toolbar.append(
    tabs,
    spacer,
    searchButton,
    copyMenu.wrap,
    exportMenu.wrap,
    structureMenu.wrap,
    // The third button keeps its place in the row and simply does not exist without a worker bridge.
    ...(agentMenu ? [agentMenu.wrap] : []),
    undoButton,
    redoButton,
    stateLabel
  )

  const formulaBar = el('div', 'wb-formula')
  const address = el('div', 'wb-address', 'A1')
  const formulaInput = el('input', 'wb-formula-input')
  formulaInput.type = 'text'
  formulaInput.spellcheck = false
  formulaInput.setAttribute('aria-label', 'Cell contents')
  const formulaKind = el('span', 'wb-formula-kind')
  formulaBar.append(address, formulaKind, formulaInput)

  const searchPanel = el('form', 'wb-search')
  searchPanel.id = 'wb-search-panel'
  searchPanel.hidden = true
  searchPanel.setAttribute('role', 'search')
  searchPanel.setAttribute('aria-label', 'Поиск по книге')
  const searchInput = el('input', 'wb-search-input')
  searchInput.type = 'search'
  searchInput.placeholder = 'Значение или формула'
  searchInput.setAttribute('aria-label', 'Найти значение или формулу в книге')
  const searchRun = el('button', 'wb-button', 'Искать')
  searchRun.type = 'submit'
  const searchPrevious = el('button', 'wb-button', 'Назад')
  const searchNext = el('button', 'wb-button', 'Далее')
  const searchMore = el('button', 'wb-button', 'Искать дальше')
  const searchClose = el('button', 'wb-button', 'Закрыть')
  for (const button of [searchPrevious, searchNext, searchMore, searchClose]) button.type = 'button'
  searchPrevious.setAttribute('aria-label', 'Предыдущее совпадение на этой странице')
  searchNext.setAttribute('aria-label', 'Следующее совпадение на этой странице')
  const searchStatus = el('span', 'wb-search-status')
  searchStatus.setAttribute('role', 'status')
  searchStatus.setAttribute('aria-live', 'polite')
  searchPanel.append(searchInput, searchRun, searchPrevious, searchNext, searchMore, searchClose, searchStatus)


  const gridHost = el('div', 'wb-grid')
  const status = el('div', 'wb-status')
  const statusAddress = el('span', 'wb-status-item')
  const statusStats = el('span', 'wb-status-item')
  const statusSheet = el('span', 'wb-status-item wb-status-right')
  status.append(statusAddress, statusStats, statusSheet)

  const notice = el('div', 'wb-notice')
  function renderNotice() {
    const notes = Array.isArray(model.notes) ? model.notes.filter(Boolean) : []
    notice.hidden = !readOnly && notes.length === 0
    notice.textContent = [
      ...(readOnly ? ['Read-only: this file was opened without edit access, so changes are off.'] : []),
      ...notes
    ].join(' ')
  }
  renderNotice()

  shell.append(toolbar, notice, formulaBar, searchPanel, gridHost, status)
  root.append(shell)

  let active = { sheet: model.active ?? 0, ref: 'A1' }
  let editing = false
  let blocked = false
  let flashTimer = null

  const grid = createGrid(gridHost, { readOnly })
  let revision = 0
  let searchPage = null
  let searchPosition = -1
  let searchPageNumber = 0

  function updateSearchStatus(message) {
    searchStatus.textContent = message
    searchPrevious.disabled = !searchPage?.matches.length
    searchNext.disabled = !searchPage?.matches.length
    searchMore.disabled = !searchPage?.cursor
  }

  function invalidateSearch() {
    revision++
    searchPage = null
    searchPosition = -1
    searchPageNumber = 0
    updateSearchStatus(searchInput.value ? 'Данные изменились. Начните поиск заново.' : 'Введите запрос и нажмите Enter.')
  }

  function noteEdit() {
    invalidateSearch()
    grid.setFormulaResultsStale(true)
  }

  function showSearchPage(result, pageNumber) {
    searchPage = result
    searchPosition = -1
    searchPageNumber = pageNumber
    if (result.reason) {
      const messages = {
        'empty-query': 'Введите запрос.',
        'query-too-long': 'Запрос слишком длинный (максимум 256 символов).',
        'cursor-stale': 'Данные изменились. Начните поиск заново.',
        'invalid-cursor': 'Продолжение недоступно. Начните поиск заново.',
        'revision-required': 'Продолжение недоступно. Начните поиск заново.'
      }
      searchPage = null
      updateSearchStatus(messages[result.reason] ?? 'Не удалось выполнить поиск.')
      return
    }
    const suffix = result.truncated ? ' Поиск не завершён; нажмите «Искать дальше».' : ' Поиск завершён.'
    updateSearchStatus(`Страница ${pageNumber}: совпадений ${result.matches.length}. Навигация только по этой странице.${suffix}`)
  }

  function runSearch() {
    try {
      showSearchPage(searchWorkbook(model, searchInput.value, { revision }), 1)
      if (searchPage?.matches.length) navigateSearch(1)
    } catch (error) {
      searchPage = null
      updateSearchStatus(`Ошибка поиска: ${describeError(error)}`)
    }
  }

  function navigateSearch(direction) {
    const matches = searchPage?.matches
    if (!matches?.length) return
    searchPosition = searchPosition < 0
      ? (direction > 0 ? 0 : matches.length - 1)
      : (searchPosition + direction + matches.length) % matches.length
    const match = matches[searchPosition]
    const index = model.sheets.findIndex((sheet) => sheet.index === match.sheetIndex && sheet.name === match.sheetName)
    if (index < 0) {
      searchPage = null
      updateSearchStatus('Лист изменился. Начните поиск заново.')
      return
    }
    const sheet = model.sheets[index]
    if (sheet.hidden?.rows?.has(match.row) || sheet.hidden?.cols?.has(match.col) ||
        sheet.rowHeights?.get(match.row) === 0 || sheet.colWidths?.get(match.col) === 0) {
      updateSearchStatus(`Совпадение ${match.sheetName}!${match.ref} находится в скрытой строке или столбце; переход невозможен. Используйте навигацию по странице.`)
      return
    }
    selectSheet(index)
    grid.setSelection({ r1: match.row, c1: match.col, r2: match.row, c2: match.col, r: match.row, c: match.col })
    const selected = grid.getSelection()
    if (selected.r !== match.row || selected.c !== match.col) {
      updateSearchStatus(`Совпадение ${match.sheetName}!${match.ref} невозможно выбрать; сетка выбрала другую ячейку.`)
      return
    }
    const suffix = searchPage.truncated ? ' Поиск не завершён; доступно продолжение.' : ' Поиск завершён.'
    updateSearchStatus(`Страница ${searchPageNumber}: ${searchPosition + 1} из ${matches.length}, ${match.sheetName}!${match.ref}. Навигация только по этой странице.${suffix}`)
  }

  function openSearch() {
    closeMenus()
    searchPanel.hidden = false
    searchButton.setAttribute('aria-expanded', 'true')
    searchInput.focus()
    searchInput.select()
  }

  function closeSearch() {
    searchPanel.hidden = true
    searchButton.setAttribute('aria-expanded', 'false')
    grid.focus?.()
  }

  searchButton.addEventListener('click', openSearch)
  searchClose.addEventListener('click', closeSearch)
  searchInput.addEventListener('input', () => {
    searchPage = null
    searchPosition = -1
    updateSearchStatus('Нажмите Enter, чтобы начать поиск.')
  })
  searchPanel.addEventListener('submit', (event) => {
    event.preventDefault()
    runSearch()
  })
  searchPanel.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      closeSearch()
    } else if (event.key === 'Enter' && event.target === searchInput && event.shiftKey && searchPage) {
      event.preventDefault()
      navigateSearch(-1)
    }
  })
  searchPrevious.addEventListener('click', () => navigateSearch(-1))
  searchNext.addEventListener('click', () => navigateSearch(1))
  searchMore.addEventListener('click', () => {
    if (!searchPage?.cursor) return
    try {
      showSearchPage(continueSearch(searchPage.cursor, { revision }), searchPageNumber + 1)
      if (searchPage?.matches.length) navigateSearch(1)
    } catch (error) {
      searchPage = null
      updateSearchStatus(`Ошибка поиска: ${describeError(error)}`)
    }
  })
  updateSearchStatus('Введите запрос и нажмите Enter.')

  const sheetLabel = () => model.sheets[active.sheet]?.name ?? ''

  function renderTabs() {
    tabs.textContent = ''
    model.sheets.forEach((sheet, index) => {
      const tab = el('button', 'wb-tab', sheet.name)
      tab.dataset.index = String(index)
      if (index === active.sheet) tab.classList.add('is-active')
      tab.addEventListener('click', () => selectSheet(index))
      tabs.append(tab)
    })
  }

  function cellAt(sheetIndex, ref) {
    const sheet = model.sheets[sheetIndex]
    if (!sheet) return null
    const parsed = parseRef(ref)
    if (!parsed) return null
    return sheet.cells.get(`${parsed.r},${parsed.c}`) ?? null
  }

  function renderFormula() {
    const cell = cellAt(active.sheet, active.ref)
    const raw = cell?.value?.formula ? `=${cell.value.formula}` : rawTextOf(cell?.value)
    if (!editing) formulaInput.value = raw
    formulaInput.readOnly = readOnly
    const size = selectionSize()
    address.textContent = size > 1 ? `${active.ref}:${endRef()}` : active.ref
    formulaKind.textContent = cell?.value?.formula ? 'fx' : ''
  }

  function selectionSize() {
    const sel = grid.getSelection?.()
    if (!sel) return 1
    return Math.max(1, (Math.abs(sel.r2 - sel.r1) + 1) * (Math.abs(sel.c2 - sel.c1) + 1))
  }

  function endRef() {
    const sel = grid.getSelection?.()
    return sel ? formatRef({ r: sel.r2, c: sel.c2 }) : active.ref
  }

  function renderStatus() {
    statusAddress.textContent = active.ref
    const sheet = model.sheets[active.sheet]
    const sel = grid.getSelection?.()
    let stats = ''
    const MAX_SCAN = 100_000
    const area = sel ? (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1) : 0
    if (sheet && sel && area > MAX_SCAN) {
      // Statistics over a million cells cost more than a keystroke may spend; say so instead.
      statusStats.textContent = `${area.toLocaleString('en-US')} cells selected`
      statusSheet.textContent = `${sheetLabel()} · ${sheet.colCount} cols × ${sheet.rowCount} rows`
      return
    }
    if (sheet && sel) {
      let count = 0
      let sum = 0
      let numeric = 0
      const r1 = Math.min(sel.r1, sel.r2)
      const r2 = Math.max(sel.r1, sel.r2)
      const c1 = Math.min(sel.c1, sel.c2)
      const c2 = Math.max(sel.c1, sel.c2)
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const cell = sheet.cells.get(`${r},${c}`)
          if (!cell) continue
          const value = cell.value
          if (value.kind === 'blank') continue
          count++
          if (value.kind === 'number' && typeof value.value === 'number') {
            sum += value.value
            numeric++
          }
        }
      }
      if (count > 0) {
        stats = `Count ${count}`
        if (numeric > 1) stats += ` · Sum ${formatNumber(sum)} · Avg ${formatNumber(sum / numeric)}`
      }
    }
    statusStats.textContent = stats
    statusSheet.textContent = `${sheetLabel()} · ${sheet ? `${sheet.colCount} cols × ${sheet.rowCount} rows` : ''}`
  }

  function renderState() {
    if (blocked) {
      undoButton.disabled = true
      redoButton.disabled = true
      return
    }
    const state = editState()
    undoButton.disabled = !state.canUndo || readOnly
    redoButton.disabled = !state.canRedo || readOnly
    if (stateLabel.classList.contains('is-error')) return
    stateLabel.textContent = state.dirty ? 'Unsaved changes' : ''
  }

  // A short toolbar note; the status bar below stays on the selection stats. An error stays longer
  // and is coloured, because it says an edit will not be saved.
  function flash(text, isError = false) {
    stateLabel.textContent = text
    stateLabel.classList.toggle('is-error', isError)
    clearTimeout(flashTimer)
    flashTimer = setTimeout(() => {
      flashTimer = null
      stateLabel.classList.remove('is-error')
      renderState()
    }, isError ? 6000 : 1600)
  }

  function selectionBox() {
    const sel = grid.getSelection?.()
    if (!sel) return null
    return { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 }
  }

  function gridHasFocus() {
    const focused = window.document.activeElement
    return Boolean(focused) && gridHost.contains(focused)
  }

  function copySelection(format) {
    const sheet = model.sheets[active.sheet]
    if (!sheet) return
    const text = format(sheet, selectionBox(), { formulas: copyFormulas })
    flash(copyText(text) ? 'Скопировано' : 'Не удалось скопировать')
  }

  function copyFormattedSelection() {
    const sheet = model.sheets[active.sheet]
    const box = selectionBox()
    if (!sheet || !box) {
      flash('Нечего копировать', true)
      return
    }
    const count = (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1)
    if (count > EXPORT_CELL_LIMIT || !selectionFitsTextBudget(sheet, box, {
      formulas: copyFormulas,
      maxChars: HANDOFF_COPY_CHAR_LIMIT
    })) {
      flash('Выделение слишком велико для безопасного копирования. Уменьшите выделение; ничего не скопировано.', true)
      return
    }
    let html = null
    let reason = ''
    try {
      html = selectionToHtml(sheet, box, { formulas: copyFormulas })
    } catch (error) {
      if (!(error instanceof RangeError)) {
        flash(`Не удалось скопировать: ${describeError(error)}`, true)
        return
      }
      reason = describeError(error)
    }
    const text = selectionToTsv(sheet, box, { formulas: copyFormulas })
    if (html !== null && copyHtml(html, text)) {
      flash('Скопировано с форматированием')
    } else {
      const copied = copyText(text)
      const detail = reason ? ` (${reason})` : ''
      flash(copied ? `Форматирование недоступно${detail}; скопирован обычный текст` : 'Не удалось скопировать ни форматированный, ни обычный текст', true)
    }
  }

  // Export takes the whole sheet, not the selection: one file per click, named after the sheet.
  function exportSheet(kind) {
    const sheet = model.sheets[active.sheet]
    if (!sheet) return
    if (sheet.rowCount * sheet.colCount > EXPORT_CELL_LIMIT) {
      flash(`This sheet is too large to export (${sheet.rowCount} × ${sheet.colCount})`, true)
      return
    }
    const files = globalThis.orca?.files
    if (typeof files?.saveAs !== 'function') {
      flash('Экспорт недоступен')
      return
    }
    const box = { r1: 0, c1: 0, r2: Math.max(0, sheet.rowCount - 1), c2: Math.max(0, sheet.colCount - 1) }
    const markdown = kind === 'markdown'
    const data = markdown
      ? selectionToMarkdown(sheet, box, { formulas: copyFormulas })
      : selectionToCsv(sheet, box, { formulas: copyFormulas })
    const name = String(sheet.name || 'sheet').replace(/[\\/:*?"<>|]+/g, '_')
    const done = (result) => {
      if (result && result.status === 'cancelled') return
      flash('Экспортировано')
    }
    try {
      const pending = files.saveAs({
        suggestedName: `${name}.${markdown ? 'md' : 'csv'}`,
        mimeType: markdown ? 'text/markdown' : 'text/csv',
        data
      })
      if (pending && typeof pending.then === 'function') {
        pending.then(done, () => flash('Экспорт не выполнен'))
      } else {
        done(pending)
      }
    } catch {
      flash('Экспорт не выполнен')
    }
  }

  function setActive(sheetIndex, ref) {
    active = { sheet: sheetIndex, ref }
    renderFormula()
    renderStatus()
    if (agentMenu?.isOpen()) renderHandoffPreview()
  }

  function selectSheet(index) {
    if (index === active.sheet || !model.sheets[index]) return
    active = { sheet: index, ref: 'A1' }
    renderTabs()
    grid.setSheet(index)
    grid.setSelection({ r1: 0, c1: 0, r2: 0, c2: 0, r: 0, c: 0 })
    setActive(index, 'A1')
    renderState()
  }

  /** Adopt a freshly loaded model and point every part of the chrome at it. */
  function adoptModel(nextModel, { resetStale }) {
    invalidateSearch()
    model = nextModel
    handoffCache = null
    // The grid switches to the new book's active sheet, so the chrome has to follow: otherwise the
    // tab, the formula bar and every later edit stay on a sheet the person is not looking at.
    const index = Math.min(Math.max(0, Number(nextModel.active) || 0), Math.max(0, nextModel.sheets.length - 1))
    active = { sheet: index, ref: active.ref }
    grid.setModel?.(nextModel)
    grid.setSheet?.(index)
    const selected = grid.getSelection()
    // The old rectangle may name cells the new sheet hides or no longer has. Pin the whole
    // selection to the clamped cell so copy and handoff cannot read a different range than the
    // formula bar shows.
    grid.setSelection({ r1: selected.r, c1: selected.c, r2: selected.r, c2: selected.c, r: selected.r, c: selected.c })
    active.ref = formatRef({ r: selected.r, c: selected.c })
    // A structural edit moves formulas and cached results, so the fresh model is not a fresh
    // calculation; an external replacement is a new workbook and starts clean.
    grid.setFormulaResultsStale(!resetStale)
    renderTabs()
    renderFormula()
    renderStatus()
    renderNotice()
    renderState()
    if (agentMenu?.isOpen()) renderHandoffPreview()
  }

  /** Rebuild the view from the staged package after a structural operation or its history move. */
  async function reloadStaged() {
    try {
      const next = await loadModel(currentDocumentBytes())
      setDateSystem(next.date1904)
      adoptModel(next, { resetStale: false })
    } catch (error) {
      flash(`Не удалось перечитать книгу: ${describeError(error)}`, true)
    }
  }

  /** Insert or delete whole rows/columns for the selected range. */
  async function structuralChange(action, axis) {
    if (blocked || readOnly) return
    const selection = grid.getSelection()
    if (!selection) return
    const index = (axis === 'row' ? Math.min(selection.r1, selection.r2) : Math.min(selection.c1, selection.c2)) + 1
    const count = (axis === 'row' ? Math.abs(selection.r2 - selection.r1) : Math.abs(selection.c2 - selection.c1)) + 1
    const result = commitStructural(docApi, { sheet: sheetLabel(), axis, action, index, count })
    if (!result.ok) {
      flash(`Не удалось изменить структуру: ${result.reason}`, true)
      return
    }
    await reloadStaged()
    flash(action === 'insert' ? 'Вставлено' : 'Удалено')
  }

  /** Whether the cell is a merged block's non-anchor, which spreadsheets do not let you edit. */
  function inMergedBlock(sheetIndex, ref) {
    const sheet = model.sheets[sheetIndex]
    if (!sheet?.merges?.length) return false
    const parsed = parseRef(ref)
    if (!parsed) return false
    return sheet.merges.some(
      (m) => parsed.r >= m.r1 && parsed.r <= m.r2 && parsed.c >= m.c1 && parsed.c <= m.c2 &&
        !(parsed.r === m.r1 && parsed.c === m.c1)
    )
  }

  /** Writes one parsed value into the model and asks the session to save it. */
  function writeCell(sheetIndex, ref, after) {
    if (blocked) return false
    if (inMergedBlock(sheetIndex, ref)) {
      flash(`Cannot change ${ref}: it is inside a merged range`, true)
      return false
    }
    const before = cellAt(sheetIndex, ref)?.value ?? { kind: 'blank', value: null, display: '', formula: null }
    const result = commitEdit(docApi, { type: 'cell', sheet: model.sheets[sheetIndex].name, ref, before, after })
    if (!result.ok) {
      flash(`Cannot change ${ref}: ${result.reason}`, true)
      return false
    }
    const sheet = model.sheets[sheetIndex]
    const parsed = parseRef(ref)
    sheet.cells.set(`${parsed.r},${parsed.c}`, {
      ref,
      row: parsed.r,
      col: parsed.c,
      value: result.op.after,
      css: cellAt(sheetIndex, ref)?.css ?? { font: {}, fill: null, border: {}, align: {} }
    })
    grid.refreshCell?.(ref)
    noteEdit()
    handoffCache = null
    if (agentMenu?.isOpen()) renderHandoffPreview()
    return true
  }

  /** Clears every cell of the selection, the way Delete does in a spreadsheet. */
  function clearSelection() {
    if (readOnly || blocked) return
    const sel = grid.getSelection?.()
    if (!sel) return
    const blank = { kind: 'blank', value: null, display: '', formula: null }
    let changed = false
    for (let r = sel.r1; r <= sel.r2; r++) {
      for (let c = sel.c1; c <= sel.c2; c++) {
        const ref = formatRef({ r, c })
        const before = cellAt(active.sheet, ref)?.value
        if (!before || before.kind === 'blank') continue
        if (writeCell(active.sheet, ref, blank)) changed = true
      }
    }
    if (changed) renderFormula()
    renderState()
  }

  /** Pastes tab- and newline-separated text, anchored at the top-left of the selection. */
  function pasteText(text) {
    if (readOnly || blocked || !text) return
    const sel = grid.getSelection?.()
    const anchor = sel ? { r: sel.r1, c: sel.c1 } : parseRef(active.ref)
    if (!anchor) return
    const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n')
    let last = null
    rows.forEach((line, rowOffset) => {
      line.split('\t').forEach((field, colOffset) => {
        const ref = formatRef({ r: anchor.r + rowOffset, c: anchor.c + colOffset })
        const before = cellAt(active.sheet, ref)?.value ?? { kind: 'blank', value: null, display: '', formula: null }
        if (writeCell(active.sheet, ref, parseInput(field, before))) last = ref
      })
    })
    if (last) setActive(active.sheet, last)
    renderState()
  }

  function commitFromInput() {
    if (readOnly || blocked) return
    const ref = active.ref
    const before = cellAt(active.sheet, ref)?.value ?? { kind: 'blank', value: null, display: '' }
    const text = formulaInput.value
    const after = parseInput(text, before)
    editing = false
    const result = commitEdit(docApi, { type: 'cell', sheet: sheetLabel(), ref, before, after })
    if (!result.ok) {
      // The writer refused this cell (a merged range, a protected sheet, a value it cannot
      // express). Say so instead of leaving the grid showing a value that will never be saved.
      flash(`Cannot change ${ref}: ${result.reason}`, true)
      renderFormula()
      return
    }
    // A coalesced burst can settle back on the value the cell already had; the log reports that
    // settled value, and the grid must show it rather than the keystroke that was just dropped.
    const settled = result.op.after
    const sheet = model.sheets[active.sheet]
    const parsed = parseRef(ref)
    sheet.cells.set(`${parsed.r},${parsed.c}`, {
      ref,
      row: parsed.r,
      col: parsed.c,
      value: settled,
      css: cellAt(active.sheet, ref)?.css ?? { font: {}, fill: null, border: {}, align: {} }
    })
    grid.refreshCell?.(ref)
    noteEdit()
    handoffCache = null
    if (agentMenu?.isOpen()) renderHandoffPreview()
    renderFormula()
    renderState()
  }

  grid.on('escape', () => {
    formulaInput.focus()
    formulaInput.select()
  })

  grid.on('selection', (payload) => {
    if (!payload) return
    const sel = payload.selection ?? payload
    setActive(active.sheet, formatRef({ r: sel.r ?? sel.r1 ?? 0, c: sel.c ?? sel.c1 ?? 0 }))
  })

  grid.on('editStart', (payload) => {
    if (readOnly) return
    const ref = payload?.ref ?? active.ref
    editing = true
    setActive(active.sheet, ref)
    // Typing a character starts the edit with that character, like a spreadsheet; F2 or a
    // double-click starts with the cell's own contents.
    if (payload?.key) {
      formulaInput.value = String(payload.key)
      formulaInput.focus()
      // The caret goes after the seeded character: selecting it would make the next keystroke
      // replace it, and the first character of a typed value would be lost.
      formulaInput.setSelectionRange(formulaInput.value.length, formulaInput.value.length)
    } else {
      formulaInput.focus()
      formulaInput.select()
    }
  })

  formulaInput.addEventListener('input', () => {
    editing = true
  })
  formulaInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      commitFromInput()
      grid.focus?.()
      grid.moveSelection?.('down')
    } else if (event.key === 'Escape') {
      editing = false
      renderFormula()
      grid.focus?.()
    }
  })
  formulaInput.addEventListener('blur', () => {
    if (editing) commitFromInput()
  })

  undoButton.addEventListener('click', () => {
    const op = undo(docApi)
    if (!op) return
    refreshAfterHistory(op, 'undo')
  })
  redoButton.addEventListener('click', () => {
    const op = redo(docApi)
    if (!op) return
    refreshAfterHistory(op, 'redo')
  })

  // Orca keeps no undo for custom editors: the frame gets the shortcut and applies it itself.
  window.addEventListener('keydown', (event) => {
    const mod = event.ctrlKey || event.metaKey
    if (!mod && gridHasFocus()) {
      if (event.key === 'Delete' || event.key === 'Backspace') {
        event.preventDefault()
        clearSelection()
        return
      }
    }
    if (!mod) return
    const key = event.key.toLowerCase()
    if (key === 'f' && !event.altKey) {
      event.preventDefault()
      openSearch()
      return
    }
    if (searchPanel.contains(event.target)) return
    if (key === 'c' && !event.shiftKey && !event.altKey && gridHasFocus()) {
      event.preventDefault()
      copySelection(selectionToTsv)
      return
    }
    if (key === 'z') {
      event.preventDefault()
      const op = event.shiftKey ? redo(docApi) : undo(docApi)
      if (op) refreshAfterHistory(op, event.shiftKey ? 'redo' : 'undo')
    } else if (key === 'y') {
      event.preventDefault()
      const op = redo(docApi)
      if (op) refreshAfterHistory(op, 'redo')
    }
  })

  window.addEventListener('paste', (event) => {
    if (!gridHasFocus() || readOnly) return
    const text = event.clipboardData?.getData('text/plain')
    if (!text) return
    event.preventDefault()
    pasteText(text)
  })

  function refreshAfterHistory(op, direction) {
    // A structural record carries package snapshots, not one cell: addresses, merges and formulas
    // all moved, so the only honest view is the staged bytes themselves.
    if (op.type !== 'cell') {
      void reloadStaged()
      return
    }
    const sheet = model.sheets.find((entry) => entry.name === op.sheet)
    if (!sheet) return
    noteEdit()
    handoffCache = null
    // History can reach a different sheet; its model still needs the new value.
    const next = direction === 'undo' ? op.before : op.after
    const parsed = parseRef(op.ref)
    const existing = sheet.cells.get(`${parsed.r},${parsed.c}`)
    sheet.cells.set(`${parsed.r},${parsed.c}`, {
      ref: op.ref,
      row: parsed.r,
      col: parsed.c,
      value: next,
      css: existing?.css ?? { font: {}, fill: null, border: {}, align: {} }
    })
    if (agentMenu?.isOpen()) renderHandoffPreview()
    if (op.sheet !== sheetLabel()) {
      renderFormula()
      renderState()
      return
    }
    grid.refreshCell?.(op.ref)
    renderFormula()
    renderState()
  }

  grid.setModel?.(model)
  grid.setSelection({ r1: 0, c1: 0, r2: 0, c2: 0, r: 0, c: 0 })
  renderTabs()
  renderFormula()
  renderStatus()
  renderState()

  return {
    destroy() {
      closeMenus()
      clearTimeout(flashTimer)
      window.document.removeEventListener('pointerdown', onOutsidePointer)
      window.document.removeEventListener('keydown', onMenuEscape)
      grid.destroy?.()
    },
    setModel(nextModel) {
      // External replacement: a newly loaded workbook starts a fresh calculation lifecycle.
      adoptModel(nextModel, { resetStale: true })
    },
    onExternalChange() {
      renderState()
    },
    onSaved() {
      renderState()
    },
    flash(text, isError = false) {
      flash(text, isError)
    },
    block(reason) {
      blocked = true
      formulaInput.readOnly = true
      flash(reason, true)
      renderState()
    },
    onDiscarded() {
      invalidateSearch()
      flash('The file changed on disk; the outside version is shown', true)
      handoffCache = null
      renderTabs()
      renderFormula()
      renderStatus()
      if (agentMenu?.isOpen()) renderHandoffPreview()
    }
  }
}

/** A host or bridge failure as the one sentence a flash can show. */
function describeError(error) {
  return error instanceof Error && error.message ? error.message : String(error)
}

// A1 addressing, duplicated from the edit layer on purpose: the UI must not depend on its internals.
function parseRef(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref.toUpperCase())
  if (!match) return null
  let column = 0
  for (const char of match[1]) column = column * 26 + (char.charCodeAt(0) - 64)
  return { r: Number(match[2]) - 1, c: column - 1 }
}

function formatRef({ r, c }) {
  let name = ''
  let n = c + 1
  while (n > 0) {
    const rem = (n - 1) % 26
    name = String.fromCharCode(65 + rem) + name
    n = Math.floor((n - 1) / 26)
  }
  return `${name}${r + 1}`
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return String(value)
  const rounded = Math.round(value * 100) / 100
  return rounded.toLocaleString('en-US', { maximumFractionDigits: 2 })
}

/** What the formula bar shows for a cell that is not a formula. */
function rawTextOf(value) {
  if (!value) return ''
  if (value.kind === 'blank') return ''
  if (value.kind === 'date') {
    // A Date's toString carries a timezone and a locale; the file stores a day number, so the bar
    // shows the date the person sees, not a JavaScript rendering of it.
    const date = value.value instanceof Date ? value.value : new Date(value.value)
    if (Number.isNaN(date.getTime())) return String(value.display ?? '')
    const iso = date.toISOString()
    return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso.slice(0, 19).replace('T', ' ')
  }
  if (value.value == null) return String(value.display ?? '')
  // An <input type="text"> cannot hold a newline; showing the value with spaces keeps it editable
  // without silently rewriting the cell when the person only pressed Enter.
  return String(value.value).replace(/\r?\n/g, ' ')
}

/**
 * Reads a typed number the way a person writes one.
 *
 * A single comma followed by one or two digits is a decimal comma (`1,5`), because that is what a
 * Russian or German keyboard produces; groups of exactly three digits are thousands (`1,234,567`).
 * Anything else falls back to JavaScript's parser, so `1e3` and `-0.5` still work.
 */
function parseNumber(text) {
  const raw = text.trim()
  if (!/^-?[\d.,]+$/.test(raw)) return null
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(raw)) {
    const value = Number(raw.replace(/,/g, ''))
    return Number.isFinite(value) ? value : null
  }
  if (/^-?\d+,\d{1,2}$/.test(raw)) {
    const value = Number(raw.replace(',', '.'))
    return Number.isFinite(value) ? value : null
  }
  if (/,/.test(raw)) return null
  const value = Number(raw)
  return Number.isFinite(value) ? value : null
}

/** Typed text becomes a number, a boolean, a string or a formula, the way a spreadsheet reads it. */
function parseInput(text, before) {
  const raw = text.trim()
  if (raw === '') return { kind: 'blank', value: null, display: '', formula: null }
  // Opening a formula cell and pressing Enter changes nothing; without this the formula would be
  // rewritten on every visit, clearing its cached result.
  if (before?.formula && raw === `=${before.formula}`) return before
  if (raw.startsWith('=') && raw.length > 1) {
    // A formula keeps no cached result here: Excel recalculates on open, and the grid shows the
    // formula text until then.
    return { kind: 'text', value: null, display: text, formula: raw.slice(1) }
  }
  const numeric = parseNumber(raw)
  if (numeric !== null) return { kind: 'number', value: numeric, display: text, formula: null }
  if (raw.toUpperCase() === 'TRUE' || raw.toUpperCase() === 'FALSE') {
    const bool = raw.toUpperCase() === 'TRUE'
    return { kind: 'boolean', value: bool, display: bool ? 'TRUE' : 'FALSE', formula: null }
  }
  return { kind: 'text', value: text, display: text, formula: null }
}
