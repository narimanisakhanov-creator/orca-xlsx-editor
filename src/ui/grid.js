const DEFAULT_ROW_HEIGHT = 24
const DEFAULT_COL_WIDTH = 96
const HEADER_HEIGHT = 26
const ROW_HEADER_WIDTH = 48
const OVERSCAN = 8
/** Longest text a single cell paints; the model keeps the rest. */
const MAX_CELL_TEXT = 2000

/**
 * Render a SheetModel into a small, scrollable Excel-like surface.
 * The grid intentionally knows nothing about Orca or workbook bytes; it consumes only the
 * view model from docs/CONTRACT.md and emits browser-facing UI events.
 */
export function createGrid(container, options = {}) {
  if (!container || typeof container.appendChild !== 'function') {
    throw new TypeError('createGrid requires a DOM container')
  }

  const readOnly = Boolean(options.readOnly)
  const listeners = new Map()
  let model = null
  let sheet = null
  let sheetIndex = 0
  let rowSizes = []
  let colSizes = []
  let rowOffsets = [0]
  let colOffsets = [0]
  let totalRows = 0
  let totalCols = 0
  let merges = []
  let mergeOrigins = new Map()
  let selection = { r1: 0, c1: 0, r2: 0, c2: 0, r: 0, c: 0 }
  let selectionAnchor = { r: 0, c: 0 }
  let dragging = null
  let renderHandle = 0
  let destroyed = false
  let resizeObserver = null

  const root = document.createElement('div')
  root.className = 'orca-grid'
  root.setAttribute('role', 'grid')
  root.setAttribute('aria-label', 'Worksheet')

  const viewport = document.createElement('div')
  viewport.className = 'orca-grid-viewport'
  viewport.tabIndex = 0
  viewport.setAttribute('aria-label', 'Worksheet grid')

  const canvas = document.createElement('div')
  canvas.className = 'orca-grid-canvas'

  const bodyLayer = makeLayer('orca-grid-body')
  const freezeTopLayer = makeLayer('orca-grid-freeze-top')
  const freezeLeftLayer = makeLayer('orca-grid-freeze-left')
  const freezeCornerLayer = makeLayer('orca-grid-freeze-corner')
  const colScrollLayer = makeLayer('orca-grid-col-scroll')
  const colFrozenLayer = makeLayer('orca-grid-col-frozen')
  const rowScrollLayer = makeLayer('orca-grid-row-scroll')
  const rowFrozenLayer = makeLayer('orca-grid-row-frozen')
  const cornerLayer = makeLayer('orca-grid-corner')

  const address = document.createElement('span')
  address.className = 'orca-grid-address'
  address.textContent = 'A1'
  cornerLayer.append(address)

  canvas.append(
    bodyLayer,
    freezeTopLayer,
    freezeLeftLayer,
    freezeCornerLayer,
    colScrollLayer,
    colFrozenLayer,
    rowScrollLayer,
    rowFrozenLayer,
    cornerLayer
  )
  viewport.append(canvas)
  root.append(viewport)
  container.replaceChildren(root)

  const layers = {
    body: bodyLayer,
    freezeTop: freezeTopLayer,
    freezeLeft: freezeLeftLayer,
    freezeCorner: freezeCornerLayer,
    colScroll: colScrollLayer,
    colFrozen: colFrozenLayer,
    rowScroll: rowScrollLayer,
    rowFrozen: rowFrozenLayer,
    corner: cornerLayer
  }

  function makeLayer(className) {
    const layer = document.createElement('div')
    layer.className = className
    return layer
  }

  function on(event, fn) {
    if (typeof fn !== 'function') return () => {}
    if (!listeners.has(event)) listeners.set(event, new Set())
    listeners.get(event).add(fn)
    return () => listeners.get(event)?.delete(fn)
  }

  function emit(event, payload) {
    const handlers = listeners.get(event)
    if (!handlers) return
    for (const fn of handlers) fn(payload)
  }

  function setModel(nextModel) {
    model = nextModel ?? null
    const requestedIndex = Number.isInteger(model?.active) ? model.active : 0
    sheetIndex = model?.sheets?.[requestedIndex] ? requestedIndex : 0
    setSheetData(model?.sheets?.[sheetIndex] ?? null)
    scheduleRender()
  }

  function setSheet(index) {
    if (!model?.sheets?.[index]) return false
    if (sheetIndex === index && sheet) return true
    sheetIndex = index
    setSheetData(model.sheets[index])
    viewport.scrollTop = 0
    viewport.scrollLeft = 0
    emit('openSheet', { index, sheet })
    scheduleRender()
    return true
  }

  function setSheetData(nextSheet) {
    sheet = nextSheet
    root.setAttribute('aria-label', sheet?.name ? `Worksheet ${sheet.name}` : 'Worksheet grid')
    rowSizes = []
    colSizes = []
    rowOffsets = [0]
    colOffsets = [0]
    merges = []
    mergeOrigins = new Map()
    if (!sheet) {
      totalRows = 0
      totalCols = 0
      canvas.style.width = `${ROW_HEADER_WIDTH}px`
      canvas.style.height = `${HEADER_HEIGHT}px`
      address.textContent = formatSelectionAddress()
      clearLayers()
      return
    }

    totalRows = Math.max(0, Number(sheet.rowCount) || 0)
    totalCols = Math.max(0, Number(sheet.colCount) || 0)
    const hiddenRows = sheet.hidden?.rows instanceof Set ? sheet.hidden.rows : new Set()
    const hiddenCols = sheet.hidden?.cols instanceof Set ? sheet.hidden.cols : new Set()
    for (let r = 0; r < totalRows; r++) {
      const hidden = hiddenRows.has(r)
      const configured = sizeFromMap(sheet.rowHeights, r)
      rowSizes[r] = hidden ? 0 : configured ?? DEFAULT_ROW_HEIGHT
      rowOffsets[r + 1] = rowOffsets[r] + rowSizes[r]
    }
    for (let c = 0; c < totalCols; c++) {
      const hidden = hiddenCols.has(c)
      const configured = sizeFromMap(sheet.colWidths, c)
      colSizes[c] = hidden ? 0 : configured ?? DEFAULT_COL_WIDTH
      colOffsets[c + 1] = colOffsets[c] + colSizes[c]
    }
    merges = Array.isArray(sheet.merges) ? sheet.merges.map(normalizeMerge).filter(Boolean) : []
    for (const merge of merges) mergeOrigins.set(`${merge.r1},${merge.c1}`, merge)

    root.setAttribute('aria-rowcount', String(totalRows))
    root.setAttribute('aria-colcount', String(totalCols))
    canvas.style.width = `${Math.max(ROW_HEADER_WIDTH, ROW_HEADER_WIDTH + colOffsets[totalCols])}px`
    canvas.style.height = `${Math.max(HEADER_HEIGHT, HEADER_HEIGHT + rowOffsets[totalRows])}px`
    selection = clampSelection(selection)
    selectionAnchor = { r: selection.r1, c: selection.c1 }
    address.textContent = formatSelectionAddress()
    applyTransforms()
  }

  function sizeFromMap(map, index) {
    if (!(map instanceof Map)) return null
    const value = Number(map.get(index))
    return Number.isFinite(value) && value >= 0 ? value : null
  }

  function normalizeMerge(value) {
    if (!value) return null
    const r1 = clampInt(value.r1, 0, Math.max(0, totalRows - 1))
    const c1 = clampInt(value.c1, 0, Math.max(0, totalCols - 1))
    const r2 = clampInt(value.r2, r1, Math.max(0, totalRows - 1))
    const c2 = clampInt(value.c2, c1, Math.max(0, totalCols - 1))
    return { r1, c1, r2, c2 }
  }

  function clampInt(value, min, max) {
    const numeric = Number.isFinite(Number(value)) ? Math.trunc(Number(value)) : min
    return Math.max(min, Math.min(max, numeric))
  }

  function clampSelection(value) {
    const maxRow = Math.max(0, totalRows - 1)
    const maxCol = Math.max(0, totalCols - 1)
    const r = visibleIndex(rowSizes, clampInt(value?.r ?? value?.r1 ?? 0, 0, maxRow), 1)
    const c = visibleIndex(colSizes, clampInt(value?.c ?? value?.c1 ?? 0, 0, maxCol), 1)
    const r1 = clampInt(value?.r1 ?? r, 0, maxRow)
    const c1 = clampInt(value?.c1 ?? c, 0, maxCol)
    const r2 = clampInt(value?.r2 ?? r, 0, maxRow)
    const c2 = clampInt(value?.c2 ?? c, 0, maxCol)
    return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2), r, c }
  }

  function setSelection(value, options = {}) {
    if (!sheet) return
    const next = expandMerges(clampSelection(value))
    selection = next
    if (options.anchor !== false) selectionAnchor = { r: next.r, c: next.c }
    address.textContent = formatSelectionAddress()
    scheduleRender()
    scrollCellIntoView(selection.r, selection.c)
    if (options.emit !== false) emit('selection', { selection: { ...selection } })
  }

  function getSelection() {
    return { ...selection }
  }

  function visibleIndex(sizes, index, direction) {
    if (!sizes.length) return 0
    const start = clampInt(index, 0, sizes.length - 1)
    let next = start
    const step = direction < 0 ? -1 : 1
    while (next >= 0 && next < sizes.length && sizes[next] <= 0) next += step
    if (next >= 0 && next < sizes.length) return next
    next = start
    while (next >= 0 && next < sizes.length && sizes[next] <= 0) next -= step
    return next >= 0 && next < sizes.length ? next : 0
  }

  function expandMerges(value) {
    const next = { ...value }
    const activeMerge = mergeForCell(next.r, next.c)
    if (activeMerge) {
      next.r = activeMerge.r1
      next.c = activeMerge.c1
    }
    let expanded = true
    while (expanded) {
      expanded = false
      for (const merge of merges) {
        if (merge.r1 > next.r2 || merge.r2 < next.r1 || merge.c1 > next.c2 || merge.c2 < next.c1) continue
        const r1 = Math.min(next.r1, merge.r1)
        const c1 = Math.min(next.c1, merge.c1)
        const r2 = Math.max(next.r2, merge.r2)
        const c2 = Math.max(next.c2, merge.c2)
        if (r1 !== next.r1 || c1 !== next.c1 || r2 !== next.r2 || c2 !== next.c2) expanded = true
        Object.assign(next, { r1, c1, r2, c2 })
      }
    }
    return next
  }

  function refreshCell(ref) {
    if (!sheet || !ref) return
    const match = /^([A-Z]+)(\d+)$/.exec(String(ref).toUpperCase())
    if (!match) return
    let c = 0
    for (const char of match[1]) c = c * 26 + char.charCodeAt(0) - 64
    const position = { r: Number(match[2]) - 1, c: c - 1 }
    if (position.r < 0 || position.c < 0 || position.r >= totalRows || position.c >= totalCols) return
    const merge = mergeForCell(position.r, position.c)
    const origin = merge ? { r: merge.r1, c: merge.c1 } : position
    const selector = `.orca-grid-cell[data-row="${origin.r}"][data-col="${origin.c}"]`
    const current = [...root.querySelectorAll(selector)]
    if (!current.length) return
    for (const oldCell of current) {
      const nextCell = createCell(origin.r, origin.c, merge)
      if (nextCell) oldCell.replaceWith(nextCell)
    }
  }

  function moveSelection(direction) {
    if (!sheet) return
    let r = selection.r
    let c = selection.c
    const name = String(direction).toLowerCase()
    switch (name) {
      case 'up': r -= 1; break
      case 'down': r += 1; break
      case 'left': c -= 1; break
      case 'right': c += 1; break
      case 'pageup': r -= pageRows(-1); break
      case 'pagedown': r += pageRows(1); break
      case 'home': c = 0; break
      case 'end': r = totalRows - 1; c = totalCols - 1; break
      case 'top':
      case 'ctrlhome': r = 0; c = 0; break
      default: return
    }
    moveTo(r, c, false)
  }

  function focus() {
    viewport.focus()
  }

  function destroy() {
    if (destroyed) return
    destroyed = true
    viewport.removeEventListener('scroll', handleScroll)
    viewport.removeEventListener('keydown', handleKeyDown)
    viewport.removeEventListener('pointerdown', handlePointerDown)
    viewport.removeEventListener('pointermove', handlePointerMove)
    viewport.removeEventListener('pointerup', handlePointerUp)
    viewport.removeEventListener('pointercancel', handlePointerUp)
    viewport.removeEventListener('dblclick', handleDoubleClick)
    resizeObserver?.disconnect()
    if (renderHandle) cancelAnimationFrameSafe(renderHandle)
    listeners.clear()
    container.replaceChildren()
  }

  function scheduleRender() {
    if (destroyed || renderHandle) return
    renderHandle = requestAnimationFrameSafe(() => {
      renderHandle = 0
      render()
    })
  }

  function render() {
    if (destroyed) return
    if (!sheet) {
      clearLayers()
      return
    }
    root.setAttribute('aria-rowcount', String(totalRows))
    root.setAttribute('aria-colcount', String(totalCols))
    canvas.style.width = `${Math.max(ROW_HEADER_WIDTH, ROW_HEADER_WIDTH + colOffsets[totalCols])}px`
    canvas.style.height = `${Math.max(HEADER_HEIGHT, HEADER_HEIGHT + rowOffsets[totalRows])}px`
    const viewportWidth = Math.max(1, viewport.clientWidth || 900)
    const viewportHeight = Math.max(1, viewport.clientHeight || 500)
    const scrollLeft = viewport.scrollLeft
    const scrollTop = viewport.scrollTop
    const freezeRows = freezeCount(sheet.freeze?.rows, totalRows)
    const freezeCols = freezeCount(sheet.freeze?.cols, totalCols)
    const frozenHeight = rowOffsets[freezeRows]
    const frozenWidth = colOffsets[freezeCols]
    const rowRange = frozenHeight >= viewportHeight - HEADER_HEIGHT
      ? { start: totalRows, end: totalRows - 1 }
      : visibleRange(rowOffsets, scrollTop + frozenHeight, scrollTop + viewportHeight - HEADER_HEIGHT, OVERSCAN, totalRows)
    const colRange = frozenWidth >= viewportWidth - ROW_HEADER_WIDTH
      ? { start: totalCols, end: totalCols - 1 }
      : visibleRange(colOffsets, scrollLeft + frozenWidth, scrollLeft + viewportWidth - ROW_HEADER_WIDTH, OVERSCAN, totalCols)
    const frozenRowRange = visibleRange(rowOffsets, 0, Math.min(frozenHeight, viewportHeight - HEADER_HEIGHT), OVERSCAN, freezeRows)
    const frozenColRange = visibleRange(colOffsets, 0, Math.min(frozenWidth, viewportWidth - ROW_HEADER_WIDTH), OVERSCAN, freezeCols)
    const regularRows = rangeList(Math.max(freezeRows, rowRange.start), rowRange.end)
    const regularCols = rangeList(Math.max(freezeCols, colRange.start), colRange.end)
    const frozenRows = rangeList(frozenRowRange.start, frozenRowRange.end)
    const frozenCols = rangeList(frozenColRange.start, frozenColRange.end)

    setupLayer(bodyLayer, ROW_HEADER_WIDTH, HEADER_HEIGHT, colOffsets[totalCols], rowOffsets[totalRows], 1)
    setupLayer(freezeTopLayer, ROW_HEADER_WIDTH, HEADER_HEIGHT, colOffsets[totalCols], frozenHeight, 3)
    setupLayer(freezeLeftLayer, ROW_HEADER_WIDTH, HEADER_HEIGHT, frozenWidth, rowOffsets[totalRows], 3)
    setupLayer(freezeCornerLayer, ROW_HEADER_WIDTH, HEADER_HEIGHT, frozenWidth, frozenHeight, 4)
    setupLayer(colScrollLayer, ROW_HEADER_WIDTH, 0, colOffsets[totalCols], HEADER_HEIGHT, 5)
    setupLayer(colFrozenLayer, ROW_HEADER_WIDTH, 0, frozenWidth, HEADER_HEIGHT, 6)
    setupLayer(rowScrollLayer, 0, HEADER_HEIGHT, ROW_HEADER_WIDTH, rowOffsets[totalRows], 5)
    setupLayer(rowFrozenLayer, 0, HEADER_HEIGHT, ROW_HEADER_WIDTH, frozenHeight, 6)
    setupLayer(cornerLayer, 0, 0, ROW_HEADER_WIDTH, HEADER_HEIGHT, 7)

    applyTransforms()
    renderGroup(bodyLayer, rowRange, colRange, regularRows, regularCols, 'body', freezeRows, freezeCols)
    renderGroup(freezeTopLayer, frozenRowRange, colRange, frozenRows, regularCols, 'top', freezeRows, freezeCols)
    renderGroup(freezeLeftLayer, rowRange, frozenColRange, regularRows, frozenCols, 'left', freezeRows, freezeCols)
    renderGroup(freezeCornerLayer, frozenRowRange, frozenColRange, frozenRows, frozenCols, 'corner', freezeRows, freezeCols)
    renderHeaders(regularRows, regularCols, frozenRows, frozenCols)
    root.style.setProperty('--grid-frozen-height', `${frozenHeight}px`)
    root.style.setProperty('--grid-frozen-width', `${frozenWidth}px`)
    root.classList.toggle('has-frozen-rows', frozenHeight > 0)
    root.classList.toggle('has-frozen-cols', frozenWidth > 0)
    address.textContent = formatSelectionAddress()
    canvas.style.minWidth = `${ROW_HEADER_WIDTH + colOffsets[totalCols]}px`
  }

  function setupLayer(layer, left, top, width, height, zIndex) {
    layer.style.left = `${left}px`
    layer.style.top = `${top}px`
    layer.style.width = `${Math.max(0, width)}px`
    layer.style.height = `${Math.max(0, height)}px`
    layer.style.zIndex = String(zIndex)
  }

  function clearLayers() {
    for (const layer of Object.values(layers)) {
      if (layer !== cornerLayer) layer.replaceChildren()
    }
  }

  function renderGroup(layer, rowRange, colRange, rows, cols, group, freezeRows, freezeCols) {
    const fragment = document.createDocumentFragment()
    const rendered = new Set()
    for (const r of rows) {
      for (const c of cols) {
        const merge = mergeForCell(r, c)
        if (merge) {
          if (merge.r1 !== r || merge.c1 !== c || mergeGroup(merge, freezeRows, freezeCols) !== group) continue
          const key = `${merge.r1},${merge.c1}`
          if (rendered.has(key)) continue
          rendered.add(key)
          const cell = createCell(merge.r1, merge.c1, merge)
          if (cell) fragment.append(cell)
          continue
        }
        const key = `${r},${c}`
        if (rendered.has(key) || mergeGroupForCoordinate(r, c, freezeRows, freezeCols) !== group) continue
        rendered.add(key)
        const cell = createCell(r, c, null)
        if (cell) fragment.append(cell)
      }
    }

    // A merged cell may begin just outside the overscan window while its visible portion is inside.
    for (const merge of merges) {
      if (mergeGroup(merge, freezeRows, freezeCols) !== group) continue
      if (!rangesIntersect(rowRange, merge.r1, merge.r2) || !rangesIntersect(colRange, merge.c1, merge.c2)) continue
      const key = `${merge.r1},${merge.c1}`
      if (rendered.has(key)) continue
      rendered.add(key)
      const cell = createCell(merge.r1, merge.c1, merge)
      if (cell) fragment.append(cell)
    }
    layer.replaceChildren(fragment)
  }

  function renderHeaders(rows, cols, frozenRows, frozenCols) {
    const colScroll = document.createDocumentFragment()
    const colFrozen = document.createDocumentFragment()
    for (const c of cols) {
      if (colSizes[c] > 0) colScroll.append(createColumnHeader(c))
    }
    for (const c of frozenCols) {
      if (colSizes[c] > 0) colFrozen.append(createColumnHeader(c))
    }
    colScrollLayer.replaceChildren(colScroll)
    colFrozenLayer.replaceChildren(colFrozen)
    const rowScroll = document.createDocumentFragment()
    const rowFrozen = document.createDocumentFragment()
    for (const r of rows) {
      if (rowSizes[r] > 0) rowScroll.append(createRowHeader(r))
    }
    for (const r of frozenRows) {
      if (rowSizes[r] > 0) rowFrozen.append(createRowHeader(r))
    }
    rowScrollLayer.replaceChildren(rowScroll)
    rowFrozenLayer.replaceChildren(rowFrozen)
  }

  function createColumnHeader(c) {
    const header = document.createElement('div')
    header.className = 'orca-grid-header orca-grid-column-header'
    if (selection.c1 <= c && c <= selection.c2) header.classList.add('is-selected')
    header.style.left = `${colOffsets[c]}px`
    header.style.width = `${colSizes[c]}px`
    header.textContent = columnName(c)
    header.setAttribute('role', 'columnheader')
    header.setAttribute('aria-label', `Column ${columnName(c)}`)
    return header
  }

  function createRowHeader(r) {
    const header = document.createElement('div')
    header.className = 'orca-grid-header orca-grid-row-header'
    if (selection.r1 <= r && r <= selection.r2) header.classList.add('is-selected')
    header.style.top = `${rowOffsets[r]}px`
    header.style.height = `${rowSizes[r]}px`
    header.textContent = String(r + 1)
    header.setAttribute('role', 'rowheader')
    header.setAttribute('aria-label', `Row ${r + 1}`)
    return header
  }

  function createCell(r, c, merge) {
    const width = merge ? colOffsets[merge.c2 + 1] - colOffsets[merge.c1] : colSizes[c]
    const height = merge ? rowOffsets[merge.r2 + 1] - rowOffsets[merge.r1] : rowSizes[r]
    if (width <= 0 || height <= 0) return null
    const view = sheet.cells instanceof Map ? sheet.cells.get(`${r},${c}`) : null
    const cell = document.createElement('div')
    cell.className = 'orca-grid-cell'
    const lastRow = merge?.r2 ?? r
    const lastCol = merge?.c2 ?? c
    const selected = selection.r1 <= lastRow && selection.r2 >= r && selection.c1 <= lastCol && selection.c2 >= c
    if (selected) cell.classList.add('is-selected')
    if (selection.r >= r && selection.r <= lastRow && selection.c >= c && selection.c <= lastCol) cell.classList.add('is-active')
    if (view?.value?.kind === 'error') cell.classList.add('is-error')
    cell.dataset.row = String(r)
    cell.dataset.col = String(c)
    cell.dataset.ref = formatRef(r, c)
    cell.style.left = `${colOffsets[merge ? merge.c1 : c]}px`
    cell.style.top = `${rowOffsets[merge ? merge.r1 : r]}px`
    cell.style.width = `${width}px`
    cell.style.height = `${height}px`
    if (view) applyCellStyle(cell, view.css)
    if (!view?.css?.align?.horizontal && (view?.value?.kind === 'number' || view?.value?.kind === 'date')) cell.style.textAlign = 'right'
    if (view?.value?.kind === 'error') cell.style.color = 'var(--orca-danger, #b91c1c)'
    const content = document.createElement('span')
    content.className = 'orca-grid-cell-content'
    content.textContent = cellDisplay(view?.value)
    if (view?.css?.font?.vertAlign === 'superscript') content.classList.add('is-superscript')
    else if (view?.css?.font?.vertAlign === 'subscript') content.classList.add('is-subscript')
    cell.append(content)
    cell.setAttribute('role', 'gridcell')
    cell.setAttribute('aria-rowindex', String(r + 1))
    cell.setAttribute('aria-colindex', String(c + 1))
    cell.setAttribute('aria-selected', selected ? 'true' : 'false')
    cell.setAttribute('aria-label', `${formatRef(r, c)}${cell.textContent ? ` ${cell.textContent}` : ''}`)
    return cell
  }

  function applyCellStyle(cell, css) {
    const style = css ?? {}
    const font = style.font ?? {}
    const fill = typeof style.fill === 'string' ? style.fill : null
    const border = style.border ?? {}
    const align = style.align ?? {}
    if (fill) cell.style.backgroundColor = fill
    if (font.bold) cell.style.fontWeight = '700'
    if (font.italic) cell.style.fontStyle = 'italic'
    // Excel stores font sizes in points; a CSS pixel is 3/4 of a point at 96 dpi. Painting 11 pt as
    // 11 px makes every workbook a quarter smaller than Excel shows it.
    if (font.size != null && Number.isFinite(Number(font.size))) {
      cell.style.fontSize = `${Math.round((Number(font.size) * 4) / 3 * 100) / 100}px`
    }
    // A filled cell carries its own light or dark background, so the theme's foreground may become
    // unreadable on it. Excel's default text is black on a painted cell; follow that unless the
    // workbook names a colour.
    if (font.color) cell.style.color = String(font.color)
    else if (fill) cell.style.color = readableOn(fill)
    if (font.family) cell.style.fontFamily = String(font.family)
    const decorations = []
    if (font.underline) decorations.push('underline')
    if (font.strike) decorations.push('line-through')
    if (decorations.length) cell.style.textDecoration = decorations.join(' ')
    if (font.underline === 'double' || font.underline === 'doubleAccounting') cell.style.textDecorationStyle = 'double'

    const horizontal = String(align.horizontal ?? '').toLowerCase()
    if (horizontal === 'center' || horizontal === 'centercontinuous' || horizontal === 'distributed') cell.style.textAlign = 'center'
    else if (horizontal === 'right' || horizontal === 'fill') cell.style.textAlign = 'right'
    else if (horizontal === 'justify') cell.style.textAlign = 'justify'
    else cell.style.textAlign = 'left'
    const vertical = String(align.vertical ?? '').toLowerCase()
    if (vertical === 'top') cell.style.justifyContent = 'flex-start'
    else if (vertical === 'bottom') cell.style.justifyContent = 'flex-end'
    else cell.style.justifyContent = 'center'
    if (align.wrapText) {
      cell.style.whiteSpace = 'normal'
      cell.style.overflowWrap = 'anywhere'
    }
    const indent = Number(align.indent)
    if (Number.isFinite(indent) && indent > 0) cell.style.paddingLeft = `${6 + indent * 8}px`
    setBorder(cell, 'top', border.top)
    setBorder(cell, 'right', border.right)
    setBorder(cell, 'bottom', border.bottom)
    setBorder(cell, 'left', border.left)
  }

  function setBorder(cell, side, value) {
    if (!value || !value.style) return
    const styleName = String(value.style).toLowerCase()
    const cssStyle = styleName === 'dashed' || styleName === 'dotted' || styleName === 'double' ? styleName : 'solid'
    const width = styleName === 'medium' || styleName === 'thick' ? 2 : 1
    const color = value.color || 'var(--orca-border, #d6d9df)'
    cell.style[`border${capitalize(side)}`] = `${width}px ${cssStyle} ${color}`
  }

  function capitalize(value) {
    return value.slice(0, 1).toUpperCase() + value.slice(1)
  }

  function cellDisplay(value) {
    if (!value || value.kind === 'blank' || value.value == null) return ''
    if (value.display != null) return String(value.display)
    if (value.value == null) return ''
    return String(value.value)
  }

  function mergeForCell(r, c) {
    const origin = mergeOrigins.get(`${r},${c}`)
    if (origin) return origin
    for (const merge of merges) {
      if (r >= merge.r1 && r <= merge.r2 && c >= merge.c1 && c <= merge.c2) return merge
    }
    return null
  }

  function mergeGroup(merge, freezeRows, freezeCols) {
    return mergeGroupForCoordinate(merge.r1, merge.c1, freezeRows, freezeCols)
  }

  function mergeGroupForCoordinate(r, c, freezeRows, freezeCols) {
    if (r < freezeRows && c < freezeCols) return 'corner'
    if (r < freezeRows) return 'top'
    if (c < freezeCols) return 'left'
    return 'body'
  }

  function rangesIntersect(range, start, end) {
    return range.start <= end && range.end >= start
  }

  function visibleRange(offsets, start, end, overscan, count) {
    if (!count) return { start: 0, end: -1 }
    const first = Math.max(0, upperBound(offsets, Math.max(0, start)) - 1 - overscan)
    const last = Math.min(count - 1, lowerBound(offsets, Math.max(0, end)) + overscan)
    return { start: Math.min(first, count - 1), end: Math.max(first, Math.min(count - 1, last)) }
  }

  function upperBound(values, target) {
    let low = 0
    let high = values.length
    while (low < high) {
      const middle = (low + high) >>> 1
      if (values[middle] <= target) low = middle + 1
      else high = middle
    }
    return low
  }

  function lowerBound(values, target) {
    let low = 0
    let high = values.length
    while (low < high) {
      const middle = (low + high) >>> 1
      if (values[middle] < target) low = middle + 1
      else high = middle
    }
    return low
  }

  function rangeList(start, end) {
    const list = []
    for (let index = start; index <= end; index++) list.push(index)
    return list
  }

  function freezeCount(value, max) {
    return Math.max(0, Math.min(max, Number.isFinite(Number(value)) ? Math.trunc(Number(value)) : 0))
  }

  function applyTransforms() {
    const x = viewport.scrollLeft
    const y = viewport.scrollTop
    freezeTopLayer.style.transform = `translate3d(0, ${y}px, 0)`
    freezeLeftLayer.style.transform = `translate3d(${x}px, 0, 0)`
    freezeCornerLayer.style.transform = `translate3d(${x}px, ${y}px, 0)`
    colScrollLayer.style.transform = `translate3d(0, ${y}px, 0)`
    colFrozenLayer.style.transform = `translate3d(${x}px, ${y}px, 0)`
    rowScrollLayer.style.transform = `translate3d(${x}px, 0, 0)`
    rowFrozenLayer.style.transform = `translate3d(${x}px, ${y}px, 0)`
    cornerLayer.style.transform = `translate3d(${x}px, ${y}px, 0)`
  }

  function scrollCellIntoView(r, c) {
    if (!sheet || !viewport.clientWidth || !viewport.clientHeight) return
    const freezeRows = freezeCount(sheet.freeze?.rows, totalRows)
    const freezeCols = freezeCount(sheet.freeze?.cols, totalCols)
    let nextTop = viewport.scrollTop
    let nextLeft = viewport.scrollLeft
    if (r >= freezeRows) {
      const top = HEADER_HEIGHT + rowOffsets[r]
      const bottom = HEADER_HEIGHT + rowOffsets[r + 1]
      const visibleTop = viewport.scrollTop + HEADER_HEIGHT + rowOffsets[freezeRows]
      const visibleBottom = viewport.scrollTop + viewport.clientHeight
      if (top < visibleTop) nextTop = top - HEADER_HEIGHT - rowOffsets[freezeRows]
      else if (bottom > visibleBottom) nextTop = bottom - viewport.clientHeight
    }
    if (c >= freezeCols) {
      const left = ROW_HEADER_WIDTH + colOffsets[c]
      const right = ROW_HEADER_WIDTH + colOffsets[c + 1]
      const visibleLeft = viewport.scrollLeft + ROW_HEADER_WIDTH + colOffsets[freezeCols]
      const visibleRight = viewport.scrollLeft + viewport.clientWidth
      if (left < visibleLeft) nextLeft = left - ROW_HEADER_WIDTH - colOffsets[freezeCols]
      else if (right > visibleRight) nextLeft = right - viewport.clientWidth
    }
    viewport.scrollTop = Math.max(0, nextTop)
    viewport.scrollLeft = Math.max(0, nextLeft)
  }

  function formatSelectionAddress() {
    if (!totalRows || !totalCols) return 'A1'
    const first = formatRef(selection.r1, selection.c1)
    const last = formatRef(selection.r2, selection.c2)
    return first === last ? first : `${first}:${last}`
  }

  function handleScroll() {
    applyTransforms()
    scheduleRender()
    emit('scroll', { left: viewport.scrollLeft, top: viewport.scrollTop })
  }

  function handleKeyDown(event) {
    if (!sheet) return
    const key = event.key
    const mod = event.ctrlKey || event.metaKey
    const extend = event.shiftKey
    let target = null
    if (key === 'ArrowUp') target = { r: selection.r - 1, c: selection.c }
    else if (key === 'ArrowDown') target = { r: selection.r + 1, c: selection.c }
    else if (key === 'ArrowLeft') target = { r: selection.r, c: selection.c - 1 }
    else if (key === 'ArrowRight') target = { r: selection.r, c: selection.c + 1 }
    else if (key === 'PageUp') target = { r: selection.r - pageRows(-1), c: selection.c }
    else if (key === 'PageDown') target = { r: selection.r + pageRows(1), c: selection.c }
    else if (key === 'Home') target = mod ? { r: 0, c: 0 } : { r: selection.r, c: 0 }
    else if (key === 'End') target = mod ? { r: totalRows - 1, c: totalCols - 1 } : { r: selection.r, c: totalCols - 1 }
    else if (key === 'Enter') target = { r: selection.r + (event.shiftKey ? -1 : 1), c: selection.c }
    else if (key === 'Tab') target = { r: selection.r, c: selection.c + (event.shiftKey ? -1 : 1) }
    if (target) {
      event.preventDefault()
      moveTo(target.r, target.c, extend)
      return
    }
    if (key === 'Escape') {
      // Leaving the grid must be possible without a pointer: Tab moves the selection, so Escape
      // hands focus back to the editor chrome.
      event.preventDefault()
      emit('escape', {})
      return
    }
    if (key === 'F2' || (!mod && !event.altKey && key.length === 1)) {
      if (readOnly) return
      event.preventDefault()
      emit('editStart', { ref: formatRef(selection.r, selection.c), key: key === 'F2' ? '' : key })
    }
  }

  function pageRows(direction) {
    const freeze = freezeCount(sheet?.freeze?.rows, totalRows)
    const frozenHeight = rowOffsets[freeze] ?? 0
    const visibleHeight = Math.max(1, viewport.clientHeight - HEADER_HEIGHT - frozenHeight)
    const from = rowOffsets[selection.r] ?? 0
    const target = Math.max(0, from + direction * visibleHeight)
    const next = visibleIndex(rowSizes, indexAtOffset(rowOffsets, target, totalRows), direction)
    return Math.max(1, Math.abs(next - selection.r))
  }

  function moveTo(r, c, extend) {
    const rowDirection = Math.sign(r - selection.r) || 1
    const colDirection = Math.sign(c - selection.c) || 1
    const activeMerge = mergeForCell(selection.r, selection.c)
    if (activeMerge) {
      if (r > selection.r && r <= activeMerge.r2) r = activeMerge.r2 + 1
      if (c > selection.c && c <= activeMerge.c2) c = activeMerge.c2 + 1
    }
    const nextR = visibleIndex(rowSizes, r, rowDirection)
    const nextC = visibleIndex(colSizes, c, colDirection)
    if (extend) {
      const anchor = selectionAnchor
      setSelection({ r1: Math.min(anchor.r, nextR), c1: Math.min(anchor.c, nextC), r2: Math.max(anchor.r, nextR), c2: Math.max(anchor.c, nextC), r: nextR, c: nextC }, { anchor: false })
      return
    }
    selectionAnchor = { r: nextR, c: nextC }
    setSelection({ r1: nextR, c1: nextC, r2: nextR, c2: nextC, r: nextR, c: nextC }, { anchor: false })
  }

  function cellFromTarget(target) {
    const cell = target instanceof Element ? target.closest('.orca-grid-cell') : null
    if (!cell || !viewport.contains(cell)) return null
    const r = Number(cell.dataset.row)
    const c = Number(cell.dataset.col)
    return Number.isInteger(r) && Number.isInteger(c) ? { r, c } : null
  }

  function cellFromPoint(event) {
    const rect = viewport.getBoundingClientRect()
    const x = event.clientX - rect.left
    const y = event.clientY - rect.top
    if (x < ROW_HEADER_WIDTH || y < HEADER_HEIGHT) return null
    const freezeRows = freezeCount(sheet?.freeze?.rows, totalRows)
    const freezeCols = freezeCount(sheet?.freeze?.cols, totalCols)
    const localX = x - ROW_HEADER_WIDTH
    const localY = y - HEADER_HEIGHT
    const bodyX = localX < colOffsets[freezeCols] ? localX : viewport.scrollLeft + localX
    const bodyY = localY < rowOffsets[freezeRows] ? localY : viewport.scrollTop + localY
    if (bodyX < 0 || bodyY < 0 || bodyX >= colOffsets[totalCols] || bodyY >= rowOffsets[totalRows]) return null
    const r = indexAtOffset(rowOffsets, bodyY, totalRows)
    const c = indexAtOffset(colOffsets, bodyX, totalCols)
    const merge = mergeForCell(r, c)
    return merge ? { r: merge.r1, c: merge.c1 } : { r, c }
  }

  function indexAtOffset(offsets, value, count) {
    if (!count) return 0
    const point = Math.max(0, value)
    let index = Math.max(0, Math.min(count - 1, upperBound(offsets, point) - 1))
    while (index < count - 1 && offsets[index + 1] <= point) index++
    return index
  }

  function handlePointerDown(event) {
    if (event.button !== 0 || !sheet) return
    const position = cellFromTarget(event.target) ?? cellFromPoint(event)
    if (!position) return
    event.preventDefault()
    // preventDefault keeps the browser from focusing the viewport itself, so a following keystroke
    // would go nowhere and typing could never start an edit.
    viewport.focus()
    dragging = { ...position, pointerId: event.pointerId }
    selectionAnchor = { r: position.r, c: position.c }
    setSelection({ r1: position.r, c1: position.c, r2: position.r, c2: position.c, r: position.r, c: position.c })
    viewport.setPointerCapture?.(event.pointerId)
  }

  function handlePointerMove(event) {
    if (!dragging || event.pointerId !== dragging.pointerId) return
    const position = cellFromPoint(event)
    if (!position) return
    const anchor = selectionAnchor
    setSelection({ r1: Math.min(anchor.r, position.r), c1: Math.min(anchor.c, position.c), r2: Math.max(anchor.r, position.r), c2: Math.max(anchor.c, position.c), r: position.r, c: position.c }, { anchor: false })
  }

  function handlePointerUp(event) {
    if (!dragging || event.pointerId !== dragging.pointerId) return
    viewport.releasePointerCapture?.(event.pointerId)
    dragging = null
  }

  function handleDoubleClick(event) {
    if (readOnly) return
    const position = cellFromTarget(event.target) ?? cellFromPoint(event)
    if (!position) return
    setSelection({ r1: position.r, c1: position.c, r2: position.r, c2: position.c, r: position.r, c: position.c })
    emit('editStart', { ref: formatRef(position.r, position.c), key: '' })
  }

  viewport.addEventListener('scroll', handleScroll, { passive: true })
  viewport.addEventListener('keydown', handleKeyDown)
  viewport.addEventListener('pointerdown', handlePointerDown)
  viewport.addEventListener('pointermove', handlePointerMove)
  viewport.addEventListener('pointerup', handlePointerUp)
  viewport.addEventListener('pointercancel', handlePointerUp)
  viewport.addEventListener('dblclick', handleDoubleClick)
  if (typeof ResizeObserver === 'function') {
    resizeObserver = new ResizeObserver(() => scheduleRender())
    resizeObserver.observe(viewport)
  }

  setModel(options.model ?? null)

  return {
    setModel,
    setSheet,
    setSelection,
    getSelection,
    refreshCell,
    moveSelection,
    focus,
    on,
    destroy
  }
}

function columnName(index) {
  let value = ''
  let number = index + 1
  while (number > 0) {
    const remainder = (number - 1) % 26
    value = String.fromCharCode(65 + remainder) + value
    number = Math.floor((number - 1) / 26)
  }
  return value
}

function formatRef(row, col) {
  return `${columnName(col)}${row + 1}`
}

/**
 * Dark ink on a light fill, light ink on a dark one: the readable default for a painted cell.
 *
 * Uses WCAG relative luminance, with the 0.179 crossover that maximises contrast between black and
 * white text. A plain luma average picks white on Excel's standard green #00B050 (2.9:1) where black
 * reaches 6.6:1.
 */
function readableOn(color) {
  const match = /^#?([0-9a-f]{6})$/i.exec(String(color).trim())
  if (!match) return 'inherit'
  const value = Number.parseInt(match[1], 16)
  const channel = (raw) => {
    const c = raw / 255
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4
  }
  const luminance =
    0.2126 * channel((value >> 16) & 0xff) +
    0.7152 * channel((value >> 8) & 0xff) +
    0.0722 * channel(value & 0xff)
  return luminance > 0.179 ? '#111111' : '#ffffff'
}

function requestAnimationFrameSafe(fn) {
  if (typeof requestAnimationFrame === 'function') return requestAnimationFrame(fn)
  return setTimeout(fn, 0)
}

function cancelAnimationFrameSafe(handle) {
  if (typeof cancelAnimationFrame === 'function') cancelAnimationFrame(handle)
  else clearTimeout(handle)
}
