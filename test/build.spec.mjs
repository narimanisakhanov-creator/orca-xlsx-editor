import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import { fileURLToPath } from 'node:url'
import vm from 'node:vm'
import { codeOnly, findDynamicCode, patchVendor } from '../build/build.mjs'

const calls = (source) => findDynamicCode(source).map((hit) => hit.what)
const BOOT = readFileSync(fileURLToPath(new URL('../editor/boot.js', import.meta.url)), 'utf8')

// Runs editor/boot.js with a captured timer, then fires it: the checkout is missing its assets
// whenever window.__xlsxReady is still unset at that point. `written` stands in for a message the
// app put on the overlay before the timer ran.
function runBoot (ready, written) {
  let fire = null
  const errors = []
  const boot = { hidden: true, textContent: 'Loading workbook…' }
  const sandbox = {
    window: { __xlsxReady: ready, setTimeout: (fn) => { fire = fn } },
    document: { getElementById: () => boot },
    console: { error: (message) => errors.push(message) }
  }
  vm.runInNewContext(BOOT, sandbox)
  if (written !== undefined) boot.textContent = written
  fire()
  return { boot, errors }
}

test('boot.js speaks up only when the editor never started', () => {
  const started = runBoot(true)
  assert.equal(started.boot.textContent, 'Loading workbook…')
  assert.deepEqual(started.errors, [])

  const missing = runBoot(false)
  assert.equal(missing.boot.hidden, false)
  assert.equal(missing.boot.textContent, 'Editor assets are missing. Run `npm run build` in the plugin folder.')
  assert.equal(missing.errors.length, 1)

  const reported = runBoot(false, 'This editor needs Orca’s desktop app.')
  assert.equal(reported.boot.textContent, 'This editor needs Orca’s desktop app.')
  assert.deepEqual(reported.errors, [])
})

test('comments never count as dynamic code', () => {
  assert.deepEqual(calls('/* lodash probes with a return-this expression */\nvar a = 1'), [])
  assert.deepEqual(calls('// Function("return this")()\nvar a = 1'), [])
  assert.deepEqual(calls('var re = /https?:\\/\\/[^ ]*/; re.test(url)'), [])
  assert.deepEqual(calls('var t = `a ${b} c`; var u = `${d}`'), [])
})

test('a call is found whatever wraps it', () => {
  assert.equal(calls('var f = Function("return this")()').length, 1)
  assert.equal(calls(';(0,Function)(x)').length, 1)
  assert.equal(calls('new Function("a")').length, 1)
  assert.equal(calls('window.Function("a")').length, 1)
  assert.equal(calls('self.Function("a")').length, 1)
  assert.equal(calls('globalThis.Function("a")').length, 1)
})

test('identifiers that merely end in Function are not calls', () => {
  assert.deepEqual(calls('totalsRowFunction(x)'), [])
  assert.deepEqual(calls('$Function(x)'), [])
  assert.deepEqual(calls('obj.Function(x)'), [])
  assert.deepEqual(calls('var s = "Function(x)"'), [])
  assert.deepEqual(calls('var s = `Function(x)`'), [])
})

test('eval is caught with and without a receiver', () => {
  assert.equal(calls('eval("1")').length, 1)
  assert.equal(calls('window.eval("1")').length, 1)
})

test('comments are blanked in place and newlines survive', () => {
  assert.equal(codeOnly('a // x\n/* y\nz */ b'), 'a     \n    \n     b')
  assert.deepEqual(findDynamicCode('line one\n// Function("x")\nline three'), [])
})

// The upstream bundle carries the probe five times and `globalThis` six, so a patched body holds
// eleven of the latter and none of the former. Any other arithmetic means upstream moved.
const PROBE = 'var probe = Function("return this")()'
const UPSTREAM = [
  ...Array.from({ length: 5 }, () => PROBE),
  ...Array.from({ length: 6 }, () => 'var base = globalThis'),
  '"function"!=typeof e&&(e=new Function(""+e))'
].join('\n')

test('the vendor patch replaces every probe and counts what it produced', () => {
  const patched = patchVendor(UPSTREAM)
  assert.equal(patched.split('Function("return this")').length - 1, 0)
  assert.equal(patched.split('globalThis').length - 1, 6 + 5 + 1) // + the header's own mention
  assert.equal(patched.split('setImmediate expects a function').length - 1, 1)
  // Already-patched input is accepted unchanged: `from` is gone, the counts still line up.
  assert.equal(patchVendor(patched.slice(patched.indexOf('*/\n') + 3)), patched)
})

test('the vendor patch fails with the exact count when upstream moved', () => {
  const moved = UPSTREAM.replace('var base = globalThis', 'var base = 1')
  assert.throws(() => patchVendor(moved), /found 0 and 10/)
})

