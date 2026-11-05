const headerCss = {
  font: { bold: true, color: '#FFFFFF', size: 13, family: 'Arial' },
  fill: '#1F4E78',
  border: {
    top: { style: 'medium', color: '#163A5B' },
    right: { style: 'thin', color: '#8EA9C1' },
    bottom: { style: 'medium', color: '#163A5B' },
    left: { style: 'thin', color: '#8EA9C1' }
  },
  align: { horizontal: 'center', vertical: 'center', wrapText: true, indent: 0 }
}

const totalRows = 501
const cells = new Map()

function addCell(row, col, value, kind = 'text', css = {}, display = value) {
  cells.set(`${row},${col}`, {
    ref: ref(row, col),
    row,
    col,
    value: { kind, value: kind === 'blank' ? null : value, display: kind === 'blank' ? '' : String(display), formula: null },
    css: { font: {}, fill: null, border: {}, align: {}, ...css }
  })
}

function ref(row, col) {
  let name = ''
  let value = col + 1
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return `${name}${row + 1}`
}

for (let col = 0; col < 8; col++) {
  addCell(0, col, ['Product', 'Region', 'Units', 'Price', 'Growth', 'Revenue', 'Date', 'Status'][col], 'text', headerCss)
}

for (let row = 1; row < totalRows; row++) {
  const units = 80 + ((row * 37) % 920)
  const price = 12.5 + ((row * 19) % 900) / 10
  const growth = ((row % 17) - 5) / 100
  addCell(row, 0, `Item ${String(row).padStart(3, '0')}`)
  addCell(row, 1, ['North', 'South', 'East', 'West'][row % 4])
  addCell(row, 2, units, 'number', { align: { horizontal: 'right', vertical: 'center' } })
  addCell(row, 3, price, 'number', { align: { horizontal: 'right', vertical: 'center' } }, price.toFixed(2))
  addCell(row, 4, growth, 'number', { align: { horizontal: 'right', vertical: 'center' } }, `${(growth * 100).toFixed(2)}%`)
  addCell(row, 5, units * price, 'number', { align: { horizontal: 'right', vertical: 'center' } }, (units * price).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }))
  const date = new Date(Date.UTC(2025, row % 12, (row % 28) + 1))
  addCell(row, 6, date, 'date', { align: { horizontal: 'center', vertical: 'center' } }, date.toISOString().slice(0, 10))
  addCell(row, 7, row % 9 === 0 ? 'Review' : 'Ready', 'text')
}

addCell(2, 0, 'Merged quarterly report', 'text', {
  font: { bold: true, color: '#17324D', size: 14 },
  fill: '#D9EAF7',
  border: {
    top: { style: 'medium', color: '#5B9BD5' },
    right: { style: 'medium', color: '#5B9BD5' },
    bottom: { style: 'medium', color: '#5B9BD5' },
    left: { style: 'medium', color: '#5B9BD5' }
  },
  align: { horizontal: 'left', vertical: 'center', wrapText: true, indent: 1 }
})

addCell(4, 7, '#DIV/0!', 'error')
addCell(5, 7, null, 'blank')
addCell(6, 7, 'On hold', 'text', {
  font: { italic: true, underline: true, strike: true },
  border: { bottom: { style: 'dashed' }, right: { style: 'dotted' } },
  align: { horizontal: 'center', vertical: 'bottom' }
})

const firstSheet = {
  name: 'Overview',
  index: 0,
  rowCount: totalRows,
  colCount: 8,
  cells,
  colWidths: new Map([[0, 178], [1, 104], [2, 84], [3, 96], [4, 96], [5, 120], [6, 112], [7, 100]]),
  rowHeights: new Map([[0, 34], [2, 30]]),
  merges: [{ r1: 2, c1: 0, r2: 2, c2: 2 }],
  freeze: { rows: 1, cols: 1 },
  hidden: { rows: new Set([8, 9]), cols: new Set([]) }
}

function makeSheet(name, index, accent) {
  const sheetCells = new Map()
  const heading = { ...headerCss, fill: accent }
  for (let col = 0; col < 5; col++) addTo(sheetCells, 0, col, ['Month', 'Orders', 'Net sales', 'Margin', 'Notes'][col], 'text', heading)
  for (let row = 1; row < 16; row++) {
    addTo(sheetCells, row, 0, `2025-${String((row % 12) + 1).padStart(2, '0')}`)
    addTo(sheetCells, row, 1, 120 + row * 11, 'number', { align: { horizontal: 'right', vertical: 'center' } })
    addTo(sheetCells, row, 2, 7800 + row * 315, 'number', { align: { horizontal: 'right', vertical: 'center' } })
    addTo(sheetCells, row, 3, `${(18 + row / 3).toFixed(1)}%`, 'number', { align: { horizontal: 'right', vertical: 'center' } })
    addTo(sheetCells, row, 4, row % 3 === 0 ? 'Watch' : 'On track')
  }
  function addTo(target, row, col, display, kind = 'text', css = {}) {
    target.set(`${row},${col}`, { ref: ref(row, col), row, col, value: { kind, value: display, display: String(display), formula: null }, css: { font: {}, fill: null, border: {}, align: {}, ...css } })
  }
  return {
    name,
    index,
    rowCount: 16,
    colCount: 5,
    cells: sheetCells,
    colWidths: new Map([[0, 120], [1, 100], [2, 120], [3, 100], [4, 130]]),
    rowHeights: new Map([[0, 34]]),
    merges: [],
    freeze: { rows: 1, cols: 0 },
    hidden: { rows: new Set(), cols: new Set() }
  }
}

export const demoWorkbook = {
  sheets: [firstSheet, makeSheet('Metrics', 1, '#548235'), makeSheet('Archive', 2, '#7030A0')],
  active: 0
}

export default demoWorkbook
