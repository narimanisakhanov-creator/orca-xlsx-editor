#!/usr/bin/env node
/**
 * Build for the sandboxed editor frame: vendor ExcelJS, bundle src/main.js into editor/assets/,
 * then assert the invariants Orca's isolated frame enforces (CSP `default-src 'self'`):
 * no dynamic code, no absolute or network URLs, every editor file under 1 MiB.
 *
 * The build is assembled and checked in memory first; assets land in editor/assets through a
 * temporary directory and a rename, and only then does editor/index.html point at them. A failed
 * build therefore leaves the previous, consistent build on disk.
 *
 * `node build/build.mjs --check-only` (npm run check:dist) inspects what is on disk without
 * bundling: markers, the module tag, asset names, sizes, references and dynamic code.
 *
 * Node only. `orca plugin check/validate` stays with the parent process.
 */
import { build } from 'esbuild'
import { createHash } from 'node:crypto'
import { mkdir, readFile, readdir, rename, rm, stat, unlink, writeFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import vm from 'node:vm'

const ROOT = path.resolve(fileURLToPath(new URL('..', import.meta.url)))
const EDITOR = path.join(ROOT, 'editor')
const ASSETS = path.join(EDITOR, 'assets')
const HTML = path.join(EDITOR, 'index.html')
const ENTRY = path.join(ROOT, 'src/main.js')

const EXCELJS_VERSION = '4.4.0'
const EXCELJS_SOURCE = path.join(ROOT, 'node_modules/exceljs/dist/exceljs.bare.min.js')
const EXCELJS_OUT_NAME = `exceljs-${EXCELJS_VERSION}.min.js`

const MAX_EDITOR_FILE_BYTES = 1024 * 1024
// Staging directory for the rename-into-place commit, never part of the shipped plugin.
const TEMP_PREFIX = '.build-tmp-'
// editor/assets carries generated code only: javascript, stylesheets and their source maps.
const ASSET_FILE = /\.(?:js|css|map)$/i

class BuildError extends Error {}

function fail (message) {
  throw new BuildError(message)
}

function countOf (haystack, needle) {
  let count = 0
  for (let at = haystack.indexOf(needle); at !== -1; at = haystack.indexOf(needle, at + needle.length)) count += 1
  return count
}

function formatBytes (bytes) {
  if (bytes < 1024) return `${bytes} B`
  const kib = bytes / 1024
  return kib < 1024 ? `${kib.toFixed(1)} KiB` : `${(kib / 1024).toFixed(2)} MiB`
}

// The build owns the line endings of the files it writes: CRLF in index.html would otherwise survive
// a rewrite on Windows and make every generated line look changed.
function withLf (text) {
  return text.replaceAll('\r\n', '\n')
}

/* ------------------------------------------------------------------ ExcelJS -------------------- */

// Two dynamic-code sites in the vendor bundle, both dead at runtime but refused by the frame's CSP
// and by `orca plugin check`:
//  1. the setImmediate polyfill builds a function from a string for a non-function argument;
//  2. the lodash global probe evaluates a return-this expression behind a short-circuit that always
//     resolves window/self first.
// Each is replaced with the value it would have produced, so behaviour is unchanged.
//
// `expected` is how many copies of `to` the finished body must carry; the patch must also leave no
// `from` behind. That pair of counts is what distinguishes "already patched" from "upstream moved"
// once `from` is gone, and it is checked on both paths.
const EXCELJS_PATCHES = [
  {
    name: 'setImmediate polyfill',
    from: '"function"!=typeof e&&(e=new Function(""+e))',
    to: '"function"!=typeof e&&(e=function(){throw new TypeError("setImmediate expects a function")})',
    expected: 1
  },
  {
    name: 'global object probe',
    from: 'Function("return this")()',
    to: 'globalThis',
    // five probes replaced on top of the six `globalThis` references upstream already contains
    expected: 5 + 6
  }
]

// No `Function(` literal may appear below: this header lands in the vendored script, and the CSP
// scan reads every comment-free byte of it. The patches are described, not quoted.
const VENDOR_HEADER = `/*!
 * ExcelJS ${EXCELJS_VERSION} (MIT, https://github.com/exceljs/exceljs) — vendored for the Orca editor frame.
 *
 * Source: node_modules/exceljs/dist/exceljs.bare.min.js, byte-identical apart from the header above
 * and the patch below. Global entry point is window.ExcelJS.
 *
 * Patches: six spans of dynamic code construction, which the frame's CSP default-src 'self' forbids
 * and \`orca plugin check\` rejects. One is the setImmediate polyfill, which built a function object
 * from a source string for a non-function argument; it now installs a callable that throws a
 * TypeError. The other five are lodash's global probe, which evaluates a return-this expression to
 * find the global object; each is replaced with globalThis, the value that short-circuit always
 * resolved to. Behaviour is unchanged in both cases.
 *
 * Regenerated and re-patched by build/build.mjs on every build — do not edit by hand.
 */
`

export function patchVendor (source) {
  let body = source
  for (const patch of EXCELJS_PATCHES) {
    body = body.replaceAll(patch.from, patch.to)
    const stray = countOf(body, patch.from)
    const copies = countOf(body, patch.to)
    if (stray !== 0 || copies !== patch.expected) {
      fail(
        `unexpected ${path.relative(ROOT, EXCELJS_SOURCE)} (${EXCELJS_VERSION}): the ${patch.name} patch ` +
        `expects no ${JSON.stringify(patch.from)} and exactly ${patch.expected} ${JSON.stringify(patch.to)}, ` +
        `found ${stray} and ${copies} — review build/build.mjs`
      )
    }
  }
  return VENDOR_HEADER + body
}

async function vendorBundle () {
  const source = await readFile(EXCELJS_SOURCE, 'utf8')
  const text = patchVendor(source)
  return { name: EXCELJS_OUT_NAME, text, bytes: Buffer.from(text) }
}

/* ------------------------------------------------------------------ bundle --------------------- */

// ExcelJS is external to the bundle: the frame loads the vendor copy as a classic script, so the
// bundle only re-exports the global. Keeps the bundle small and the vendor copy CSP-clean.
const EXCELJS_EXPORTS = [
  'Workbook',
  'ModelContainer',
  'stream',
  'ValueType',
  'FormulaType',
  'RelationshipType',
  'DocumentType',
  'ReadingOrder',
  'ErrorValue'
]

const EXCELJS_GLOBAL_SHIM = `// generated by build/build.mjs — exceljs resolved from the vendor script, not bundled
const X = globalThis.ExcelJS
if (!X) throw new Error('[xlsx] ExcelJS vendor script did not load')
export default X
${EXCELJS_EXPORTS.map((name) => `export const ${name} = X.${name}`).join('\n')}
`

export const exceljsFromGlobal = {
  name: 'exceljs-from-global',
  setup (ctx) {
    ctx.onResolve({ filter: /^exceljs$/ }, () => ({ path: 'exceljs', namespace: 'exceljs-global' }))
    ctx.onLoad({ filter: /^exceljs$/, namespace: 'exceljs-global' }, () => ({ contents: EXCELJS_GLOBAL_SHIM, loader: 'js' }))
  }
}

async function bundleApp () {
  const result = await build({
    entryPoints: [ENTRY],
    bundle: true,
    format: 'esm',
    target: 'chrome120',
    minify: true,
    legalComments: 'none',
    metafile: true,
    write: false,
    outdir: ASSETS,
    entryNames: '[name]',
    assetNames: '[name]',
    external: ['fs', 'path', 'crypto', 'stream', 'util', 'os'],
    plugins: [exceljsFromGlobal],
    logLevel: 'warning'
  })

  const js = result.outputFiles.find((file) => file.path.endsWith('.js'))
  if (!js) fail('esbuild produced no javascript output')
  // CSS arrives through `import './ui/app.css'` in the entry; esbuild emits one stylesheet.
  const css = result.outputFiles.find((file) => file.path.endsWith('.css')) ?? null

  const jsBytes = Buffer.from(js.contents)
  const cssBytes = css ? Buffer.from(css.contents) : null
  // One hash for both artifacts: index.html then carries a single build identity to grep for.
  const hash = createHash('sha256')
    .update(jsBytes)
    .update(cssBytes ?? Buffer.alloc(0))
    .digest('hex')
    .slice(0, 10)

  return {
    js: { name: `app-${hash}.js`, bytes: jsBytes, text: jsBytes.toString('utf8') },
    css: cssBytes ? { name: `app-${hash}.css`, bytes: cssBytes, text: cssBytes.toString('utf8') } : null,
    hash
  }
}

/* ------------------------------------------------------------------ index.html ----------------- */

// Marker pairs the build owns: everything between them is replaced on every build, so no tag drifts
// from the emitted file. All three are mandatory in the template.
//   app-styles     the hashed stylesheet link, empty when the entry imports no css;
//   vendor-script  the classic ExcelJS tag, named after EXCELJS_VERSION;
//   app-script     the module tag that boots the bundle.
const BLOCKS = [
  { name: 'app-styles', pattern: blockPattern('app-styles') },
  { name: 'vendor-script', pattern: blockPattern('vendor-script') },
  { name: 'app-script', pattern: blockPattern('app-script') }
]

function blockPattern (name) {
  return new RegExp(`([ \\t]*<!-- build:${name} -->\\n)([\\s\\S]*?)([ \\t]*<!-- /build:${name} -->)`)
}

const INDENT = '    '

function renderIndex (html, app) {
  const inner = {
    'app-styles': app.css ? `${INDENT}<link rel="stylesheet" href="./assets/${app.css.name}">\n` : '',
    'vendor-script': `${INDENT}<script src="./assets/${EXCELJS_OUT_NAME}"></script>\n`,
    'app-script': `${INDENT}<script type="module" src="./assets/${app.js.name}"></script>\n`
  }

  let next = html
  for (const block of BLOCKS) {
    if (!block.pattern.test(next)) fail(`editor/index.html has no build:${block.name} marker block`)
    next = next.replace(block.pattern, (_, open, _body, close) => `${open}${inner[block.name]}${close}`)
  }

  // The written page is only loadable with a deferred module entry and the vendor script in place.
  if ((next.match(/<script type="module"/g) ?? []).length !== 1) {
    fail('editor/index.html must hold exactly one type="module" script tag')
  }
  if (!next.includes(`<script src="./assets/${EXCELJS_OUT_NAME}"></script>`)) {
    fail(`editor/index.html must load ./assets/${EXCELJS_OUT_NAME}`)
  }
  return next
}

/* ------------------------------------------------------------------ checks --------------------- */

// `eval` and the Function constructor are banned by the frame's CSP, so the scan has to be a check
// on code, not on text: the vendor header describes the patched expressions, and a string can quote
// them. Comments are blanked, and so are the contents of string, template and regex literals — a
// literal can never be a call site. Newlines survive, so reported line numbers stay true, and a
// `${…}` interpolation inside a template stays code.
const IDENT_START = /[A-Za-z_$]/
const IDENTIFIER = /[A-Za-z0-9_$]+/y
// A `/` opens a regex literal, not a division, after one of these tokens.
const REGEX_AFTER = new Set([
  '', '(', '[', '{', ',', ';', ':', '=', '!', '&', '|', '?', '+', '-', '*', '%', '^', '~', '<', '>',
  'return', 'typeof', 'instanceof', 'in', 'of', 'new', 'delete', 'void', 'throw', 'case', 'do', 'else',
  'yield', 'await'
])

function blank (text) {
  return text.replace(/[^\n]/g, ' ')
}

function skipQuoted (source, start) {
  const quote = source[start]
  let i = start + 1
  while (i < source.length) {
    const ch = source[i]
    if (ch === '\\') { i += 2; continue }
    if (ch === quote) return i + 1
    i += 1
  }
  return source.length
}

function skipRegex (source, start) {
  let i = start + 1
  let inClass = false
  while (i < source.length) {
    const ch = source[i]
    if (ch === '\\') { i += 2; continue }
    if (ch === '\n') return i
    if (inClass) {
      if (ch === ']') inClass = false
      i += 1
      continue
    }
    if (ch === '[') { inClass = true; i += 1; continue }
    if (ch === '/') {
      i += 1
      while (/[a-z]/i.test(source[i] ?? '')) i += 1
      return i
    }
    i += 1
  }
  return source.length
}

function identEnd (source, start) {
  IDENTIFIER.lastIndex = start
  const match = IDENTIFIER.exec(source)
  return match ? start + match[0].length : start + 1
}

export function codeOnly (source) {
  const out = []
  // One frame per lexical context: the root, and one per `${…}` inside a template literal.
  const stack = [{ template: false, depth: 0 }]
  let prev = ''
  let i = 0

  while (i < source.length) {
    const top = stack[stack.length - 1]
    const ch = source[i]

    if (top.template) {
      if (ch === '\\') { out.push(blank(source.slice(i, i + 2))); i += 2; continue }
      if (ch === '`') { out.push(' '); i += 1; stack.pop(); prev = '`'; continue }
      if (ch === '$' && source[i + 1] === '{') {
        out.push('  ')
        i += 2
        stack.push({ template: false, depth: 0 })
        prev = '{'
        continue
      }
      out.push(blank(ch))
      i += 1
      continue
    }

    if (ch === '/' && source[i + 1] === '/') {
      const end = source.indexOf('\n', i)
      const stop = end === -1 ? source.length : end
      out.push(blank(source.slice(i, stop)))
      i = stop
      continue
    }
    if (ch === '/' && source[i + 1] === '*') {
      const close = source.indexOf('*/', i + 2)
      const stop = close === -1 ? source.length : close + 2
      out.push(blank(source.slice(i, stop)))
      i = stop
      continue
    }
    if (ch === '"' || ch === "'") {
      const stop = skipQuoted(source, i)
      out.push(blank(source.slice(i, stop)))
      prev = ch
      i = stop
      continue
    }
    if (ch === '`') {
      out.push(' ')
      i += 1
      stack.push({ template: true, depth: 0 })
      continue
    }
    if (ch === '/' && REGEX_AFTER.has(prev)) {
      const stop = skipRegex(source, i)
      out.push(blank(source.slice(i, stop)))
      prev = 'regex'
      i = stop
      continue
    }
    if (ch === '{') top.depth += 1
    else if (ch === '}') {
      // At depth 0 inside a `${…}` the brace closes the interpolation; at the root it is just code.
      if (top.depth === 0 && stack.length > 1) {
        out.push(ch)
        i += 1
        stack.pop()
        prev = '}'
        continue
      }
      top.depth = Math.max(0, top.depth - 1)
    }
    if (IDENT_START.test(ch)) {
      const stop = identEnd(source, i)
      const word = source.slice(i, stop)
      out.push(word)
      prev = word
      i = stop
      continue
    }
    if (!/\s/.test(ch)) prev = ch
    out.push(ch)
    i += 1
  }
  return out.join('')
}

// Call sites of the two dynamic-code primitives. `eval` matches on a word boundary, so `window.eval(`
// is caught too. `Function` matches after any character that cannot end an identifier, which covers
// `Function(x)`, `new Function(x)` and `window.Function(x)`; the indirect idiom `(0,Function)(x)`
// puts the wrapper's closing paren between the identifier and the call, so `) (` is accepted there.
// A `.` prefix is allowed only for the global objects such a call reaches through, because a method
// that merely happens to be called `Function` on another receiver is not the intrinsic. `\w` and `$`
// keep `totalsRowFunction(` and `$Function(` out of the report.
const DYNAMIC_CODE = [
  { pattern: /\beval\s*\(/g, what: 'eval()' },
  { pattern: /(?:^|[^\w$.])(?:Function|(?:window|self|globalThis|global)\.Function)\s*(?:\(|\)\s*\()/g, what: 'the Function constructor' }
]

export function findDynamicCode (text) {
  const code = codeOnly(text)
  const hits = []
  for (const rule of DYNAMIC_CODE) {
    const at = code.search(rule.pattern)
    if (at === -1) continue
    hits.push({ what: rule.what, line: code.slice(0, at).split('\n').length })
  }
  return hits
}

function checkDynamicCode (sources) {
  const offences = []
  for (const source of sources) {
    for (const hit of findDynamicCode(source.text)) offences.push(`${source.label}:${hit.line} calls ${hit.what}`)
  }
  if (offences.length) fail(`CSP violation — dynamic code in emitted javascript:\n  ${offences.join('\n  ')}`)
}

function checkVendorLoads (text) {
  const sandbox = {
    window: {},
    console,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    queueMicrotask,
    TextEncoder,
    TextDecoder,
    URL,
    process
  }
  sandbox.global = sandbox
  sandbox.self = sandbox.window
  try {
    vm.runInNewContext(text, sandbox, { filename: EXCELJS_OUT_NAME })
  } catch (error) {
    fail(`${EXCELJS_OUT_NAME} does not evaluate: ${error.message}`)
  }
  const api = sandbox.window.ExcelJS
  if (!api || typeof api.Workbook !== 'function') fail(`${EXCELJS_OUT_NAME} evaluated but did not expose window.ExcelJS.Workbook`)
}

// Any quoting style: generated tags use double quotes, but a hand edit with single quotes or none
// must not slip past the "every reference resolves" check.
const HTML_REF = /\b(?:src|href)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'`=<>]+))/gi

async function checkHtmlRefs (html, planned, dropped) {
  const refs = [...html.matchAll(HTML_REF)].map((match) => match[1] ?? match[2] ?? match[3])
  if (refs.length === 0) fail('editor/index.html references nothing')
  for (const ref of refs) {
    if (ref === '' || ref.startsWith('#')) continue
    if (/^[a-z][a-z0-9+.-]*:/i.test(ref) || ref.startsWith('//')) {
      fail(`editor/index.html references a non-relative url: ${ref}`)
    }
    const target = path.resolve(EDITOR, ref)
    if (target !== EDITOR && !target.startsWith(EDITOR + path.sep)) {
      fail(`editor/index.html escapes the editor directory: ${ref}`)
    }
    const rel = path.relative(EDITOR, target).split(path.sep).join('/')
    if (planned.has(rel)) continue
    if (dropped.has(rel)) fail(`editor/index.html references a file this build deletes: ${ref}`)
    const exists = await stat(target).then((info) => info.isFile(), () => false)
    if (!exists) fail(`editor/index.html references a missing file: ${ref}`)
  }
  return refs
}

// A `<script>` without `src` is code the bundle check would otherwise never see.
const SCRIPT_TAG = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi

function inlineScripts (html, label = 'editor/index.html') {
  const sources = []
  for (const match of html.matchAll(SCRIPT_TAG)) {
    if (/\bsrc\s*=/i.test(match[1])) continue
    // Data blocks (`application/json`, `application/ld+json`) carry no code.
    if (/\btype\s*=\s*["']?application\/(?:ld\+)?json/i.test(match[1])) continue
    sources.push({ label: `${label} (inline script)`, text: match[2] })
  }
  return sources
}

function checkSizes (sizes) {
  for (const [rel, bytes] of sizes) {
    if (bytes > MAX_EDITOR_FILE_BYTES) {
      fail(`editor file over the 1 MiB limit: editor/${rel} (${formatBytes(bytes)})`)
    }
  }
}

async function walk (dir) {
  const entries = await readdir(dir, { withFileTypes: true })
  const files = []
  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) files.push(...await walk(full))
    else if (entry.isFile()) {
      files.push({
        path: full,
        rel: path.relative(EDITOR, full).split(path.sep).join('/'),
        size: (await stat(full)).size
      })
    }
  }
  return files
}

/* ------------------------------------------------------------------ commit --------------------- */

// Leftovers from a build killed between staging and commit; they are never part of the plugin.
async function sweepTemp () {
  for (const entry of await readdir(EDITOR, { withFileTypes: true })) {
    if (entry.name.startsWith(TEMP_PREFIX)) await rm(path.join(EDITOR, entry.name), { recursive: true, force: true })
  }
}

// Stage every changed file next to its target, then rename into place: a rename is the only write
// that can fail without leaving a truncated file behind. Assets first, index.html second, the files
// index.html no longer names last — at no instant does the frame see a page whose assets are absent.
async function commit (planned, html, stale) {
  const staging = path.join(EDITOR, `${TEMP_PREFIX}${process.pid.toString(36)}-${Date.now().toString(36)}`)
  await mkdir(staging, { recursive: true })
  try {
    for (const [rel, file] of planned) {
      const target = path.join(EDITOR, rel)
      const current = await readFile(target).catch(() => null)
      if (current && current.equals(file.bytes)) continue
      const staged = path.join(staging, path.posix.basename(rel))
      await writeFile(staged, file.bytes)
      await rename(staged, target)
    }
  } finally {
    await rm(staging, { recursive: true, force: true })
  }

  const current = await readFile(HTML, 'utf8').catch(() => null)
  if (current !== html) {
    const staged = path.join(EDITOR, `${TEMP_PREFIX}index.html`)
    await writeFile(staged, html)
    await rename(staged, HTML)
  }

  for (const name of stale) await unlink(path.join(ASSETS, name))
}

/* ------------------------------------------------------------------ run ------------------------ */

async function buildAll () {
  const started = Date.now()
  await mkdir(ASSETS, { recursive: true })
  await sweepTemp()

  const vendor = await vendorBundle()
  const app = await bundleApp()
  const html = renderIndex(
    withLf(await readFile(HTML, 'utf8').catch(() => fail('editor/index.html is missing — restore the frame template'))),
    app
  )

  // Everything the build writes, keyed by its path under editor/.
  const planned = new Map([
    [`assets/${vendor.name}`, { bytes: vendor.bytes, text: vendor.text }],
    [`assets/${app.js.name}`, { bytes: app.js.bytes, text: app.js.text }]
  ])
  if (app.css) planned.set(`assets/${app.css.name}`, { bytes: app.css.bytes, text: app.css.text })

  // editor/assets holds the current build and nothing else: any other script, stylesheet or map is a
  // leftover from a previous version or hash.
  const keep = new Set([...planned.keys()].map((rel) => path.posix.basename(rel)))
  const stale = (await readdir(ASSETS, { withFileTypes: true }))
    .filter((entry) => entry.isFile() && ASSET_FILE.test(entry.name) && !keep.has(entry.name))
    .map((entry) => entry.name)
  const dropped = new Set(stale.map((name) => `assets/${name}`))

  // Validate the state the commit will leave behind, not the state on disk.
  const existing = await walk(EDITOR)
  const sizes = new Map()
  for (const file of existing) if (!dropped.has(file.rel)) sizes.set(file.rel, file.size)
  for (const [rel, file] of planned) sizes.set(rel, file.bytes.length)
  sizes.set('index.html', Buffer.byteLength(html))
  checkSizes(sizes)

  const sources = [...planned]
    .filter(([rel]) => rel.endsWith('.js'))
    .map(([rel, file]) => ({ label: `editor/${rel}`, text: file.text }))
  for (const file of existing) {
    if (!file.rel.endsWith('.js') || dropped.has(file.rel) || planned.has(file.rel)) continue
    sources.push({ label: `editor/${file.rel}`, text: await readFile(file.path, 'utf8') })
  }
  sources.push(...inlineScripts(html))
  checkDynamicCode(sources)

  checkVendorLoads(vendor.text)
  const refs = await checkHtmlRefs(html, new Set(planned.keys()), dropped)

  await commit(planned, html, stale)

  console.log(`[xlsx] vendor  editor/assets/${vendor.name}  ${formatBytes(vendor.bytes.length)}  (setImmediate polyfill and 5 global probes patched)`)
  console.log(`[xlsx] bundle  editor/assets/${app.js.name}  ${formatBytes(app.js.bytes.length)}  sha256:${app.hash}`)
  if (app.css) console.log(`[xlsx] styles  editor/assets/${app.css.name}  ${formatBytes(app.css.bytes.length)}`)
  else console.log('[xlsx] styles  none (entry imports no css)')
  console.log(`[xlsx] html    editor/index.html  ${formatBytes(Buffer.byteLength(html))}  (vendor and module tags regenerated)`)

  const rows = [...sizes].sort(([a], [b]) => a.localeCompare(b))
  const total = rows.reduce((sum, [, bytes]) => sum + bytes, 0)
  console.log(`[xlsx] editor/  ${rows.length} files, ${formatBytes(total)} total, largest ${formatBytes(Math.max(...rows.map(([, bytes]) => bytes)))}`)
  for (const [rel, bytes] of rows) {
    console.log(`         editor/${rel.padEnd(34)} ${formatBytes(bytes).padStart(10)}`)
  }
  console.log(`[xlsx] checks  no eval/Function constructor outside comments and literals · vendor evaluates in vm · ${refs.length} relative refs resolve · every file ≤ 1 MiB`)
  console.log(`[xlsx] done in ${Date.now() - started} ms`)
}

async function checkDist () {
  const started = Date.now()
  const html = await readFile(HTML, 'utf8').catch(() => fail('editor/index.html is missing — run `npm run build`'))
  if (html.includes('\r')) fail('editor/index.html has CRLF line endings — run `npm run build`')

  const inner = new Map()
  for (const block of BLOCKS) {
    const match = block.pattern.exec(html)
    if (!match) fail(`editor/index.html has no build:${block.name} marker block — run \`npm run build\``)
    inner.set(block.name, match[2])
  }

  const appTag = /<script type="module" src="\.\/assets\/(app-[0-9a-f]{10}\.js)"><\/script>/.exec(inner.get('app-script'))
  if (!appTag) fail('editor/index.html: the build:app-script block is not a type="module" ./assets/app-<hash>.js tag')
  const appName = appTag[1]

  const vendorTag = /<script src="\.\/assets\/(exceljs-[^"]+)"><\/script>/.exec(inner.get('vendor-script'))
  if (!vendorTag) fail('editor/index.html: the build:vendor-script block is not a classic ./assets/exceljs-<version>.min.js tag')
  if (vendorTag[1] !== EXCELJS_OUT_NAME) {
    fail(`editor/index.html loads ${vendorTag[1]} but the build vendors ${EXCELJS_OUT_NAME} — run \`npm run build\``)
  }

  const styleTag = inner.get('app-styles').trim()
  let cssName = null
  if (styleTag !== '') {
    const match = /^<link rel="stylesheet" href="\.\/assets\/(app-[0-9a-f]{10}\.css)">$/.exec(styleTag)
    if (!match) fail('editor/index.html: the build:app-styles block is neither empty nor a single ./assets/app-<hash>.css link')
    cssName = match[1]
  }

  const expected = new Set([EXCELJS_OUT_NAME, appName, ...(cssName ? [cssName] : [])])
  if (cssName && cssName.replace(/\.css$/, '.js') !== appName) {
    fail(`editor/index.html pairs ${appName} with ${cssName}; one build emits one hash — run \`npm run build\``)
  }

  const entries = await readdir(ASSETS, { withFileTypes: true }).catch(() => fail('editor/assets is missing — run `npm run build`'))
  const names = entries.filter((entry) => entry.isFile()).map((entry) => entry.name)
  const missing = [...expected].filter((name) => !names.includes(name))
  if (missing.length) fail(`editor/assets is missing ${missing.join(', ')} — run \`npm run build\``)
  const stray = names.filter((name) => ASSET_FILE.test(name) && !expected.has(name))
  if (stray.length) fail(`editor/assets holds files index.html does not name: ${stray.join(', ')} — run \`npm run build\``)

  const files = await walk(EDITOR)
  checkSizes(new Map(files.map((file) => [file.rel, file.size])))

  const sources = []
  for (const file of files) {
    if (file.rel.endsWith('.js')) sources.push({ label: `editor/${file.rel}`, text: await readFile(file.path, 'utf8') })
  }
  sources.push(...inlineScripts(html))
  checkDynamicCode(sources)

  const refs = await checkHtmlRefs(html, new Set(), new Set())

  // index.html names one build identity; the pair on disk has to reproduce it.
  const appHash = createHash('sha256')
    .update(await readFile(path.join(ASSETS, appName)))
    .update(cssName ? await readFile(path.join(ASSETS, cssName)) : Buffer.alloc(0))
    .digest('hex')
    .slice(0, 10)
  if (`app-${appHash}.js` !== appName) {
    fail(`editor/assets/${appName} hashes to app-${appHash}.js — the pair is stale, run \`npm run build\``)
  }

  // Only the repository can prove the vendor bytes are the patched upstream; an installed plugin has
  // no node_modules and reports that the check was skipped.
  const upstream = await readFile(EXCELJS_SOURCE, 'utf8').catch(() => null)
  const vendorVerified = upstream !== null
  if (vendorVerified) {
    const actual = await readFile(path.join(ASSETS, EXCELJS_OUT_NAME), 'utf8')
    if (actual !== patchVendor(upstream)) {
      fail(`editor/assets/${EXCELJS_OUT_NAME} is not what build/build.mjs emits — run \`npm run build\``)
    }
  }

  const total = files.reduce((sum, file) => sum + file.size, 0)
  console.log(`[xlsx] check  editor/index.html + editor/assets agree (${expected.size} generated files, ${formatBytes(total)} under editor/)`)
  console.log(`[xlsx] check  ${refs.length} relative refs resolve · no dynamic code outside comments and literals · every file ≤ 1 MiB`)
  console.log(`[xlsx] check  vendor ${vendorVerified ? `matches patched ${EXCELJS_VERSION} upstream` : `not re-verified (${path.relative(ROOT, EXCELJS_SOURCE)} is absent)`}`)
  console.log(`[xlsx] check ok in ${Date.now() - started} ms`)
}

async function main (argv = process.argv.slice(2)) {
  const unknown = argv.filter((arg) => arg !== '--check-only')
  if (unknown.length) fail(`unknown argument(s): ${unknown.join(' ')} — usage: node build/build.mjs [--check-only]`)
  return argv.includes('--check-only') ? checkDist() : buildAll()
}

// Importable without side effects so tests can reuse the scanner and the exceljs resolution plugin.
const invokedAsScript = process.argv[1] != null &&
  path.resolve(process.argv[1]).toLowerCase() === path.resolve(fileURLToPath(import.meta.url)).toLowerCase()

if (invokedAsScript) {
  const checkOnly = process.argv.includes('--check-only')
  main().catch((error) => {
    console.error(`[xlsx] ${checkOnly ? 'check' : 'build'} failed: ${error.message}`)
    if (process.env.XLSX_DEBUG && !(error instanceof BuildError)) console.error(error.stack)
    process.exitCode = 1
  })
}
