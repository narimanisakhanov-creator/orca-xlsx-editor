#!/usr/bin/env node
import { build } from 'esbuild'
import { mkdir } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const root = path.resolve(fileURLToPath(new URL('..', import.meta.url)))
const outDir = path.join(root, 'demo', 'dist')
await mkdir(outDir, { recursive: true })
await build({
  stdin: {
    contents: `import { createGrid } from './src/ui/grid.js'
import workbook from './demo/demo-data.js'
import './src/ui/grid.css'

const params = new URLSearchParams(location.search)
document.documentElement.dataset.theme = params.get('theme') === 'dark' ? 'dark' : 'light'
const host = document.querySelector('#grid')
const status = document.querySelector('#status')
const sheetPicker = document.querySelector('#sheets')
const themePicker = document.querySelector('#theme')
const grid = createGrid(host, { readOnly: !params.has('editable') })
grid.setModel(workbook)
for (const [index, sheet] of workbook.sheets.entries()) {
  const option = document.createElement('option')
  option.value = String(index)
  option.textContent = sheet.name
  sheetPicker.append(option)
}
sheetPicker.addEventListener('change', () => {
  grid.setSheet(Number(sheetPicker.value))
  grid.setSelection({ r: 0, c: 0 })
  grid.focus()
})
themePicker.value = document.documentElement.dataset.theme
themePicker.addEventListener('change', () => { document.documentElement.dataset.theme = themePicker.value })
function ref(r, c) {
  let name = ''
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name
  return name + String(r + 1)
}
grid.on('selection', ({ selection }) => {
  const start = ref(selection.r1, selection.c1)
  const end = ref(selection.r2, selection.c2)
  status.textContent = (start === end ? start : start + ':' + end) + ' · ' + workbook.sheets[Number(sheetPicker.value)].name
})
grid.on('editStart', ({ ref: address, key }) => {
  status.textContent = 'Edit requested: ' + address + (key ? ' · ' + key : '')
})
export { grid, workbook, createGrid }
`,
    resolveDir: root
  },
  bundle: true,
  format: 'iife',
  globalName: 'OrcaGridDemo',
  target: 'es2020',
  minify: false,
  outfile: path.join(outDir, 'demo.js'),
  logLevel: 'info'
})
