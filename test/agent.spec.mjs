// Handing a table to an agent: the pure text rules in worker/handoff.js plus the worker handlers
// themselves, mounted on a stand-in `orca` so no terminal and no Orca window are involved.
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import path from 'node:path'
import { TEXT_LIMIT, buildHandoff, rankAgents, readSend, truncateText } from '../worker/handoff.js'
import activate from '../worker/main.js'
import { selectionToMarkdown, selectionToTsv } from '../src/ui/clipboard.js'
import { loadWorkbook } from '../src/model/workbook.js'

const HERE = path.dirname(fileURLToPath(import.meta.url))

const agent = (terminalId, state, stateStartedAt, title = null) => ({ terminalId, state, stateStartedAt, title, agent: null })

/** A stand-in host: records calls, answers from `results`, and fails where a result is an Error. */
function fakeOrca(results = {}) {
  const handlers = new Map()
  const calls = []
  return {
    handlers,
    calls,
    panels: {
      handle(name, handler) {
        handlers.set(name, handler)
        return { dispose: () => handlers.delete(name) }
      }
    },
    host: {
      call: async (method, params) => {
        calls.push({ method, params })
        const answer = results[method]
        if (answer instanceof Error) throw answer
        return typeof answer === 'function' ? answer(params) : answer
      }
    }
  }
}

test('the host text limit is the 4096 Orca enforces', () => {
  assert.equal(TEXT_LIMIT, 4096)
})

test('a handoff is one context line and the table, for a real workbook range', async () => {
  const workbook = await loadWorkbook(new Uint8Array(readFileSync(path.join(HERE, 'fixtures/styled.xlsx'))))
  const sheet = workbook.sheets[0]
  const box = { r1: 0, c1: 0, r2: 3, c2: 3 }

  const markdown = buildHandoff({
    sheetName: sheet.name,
    range: 'A1:D4',
    format: 'markdown',
    cells: selectionToMarkdown(sheet, box)
  })
  assert.equal(
    markdown,
    [
      'Таблица из XLSX: лист «Ledger», диапазон A1:D4',
      '',
      '| Store | Sales | Revenue | Plan |',
      '| --- | ---: | --- | ---: |',
      '| 4002_Store 20000000000000 | 200,274.00 | B2*48.7 | 250,200 |',
      '| 4003_Store 30000000000000 | 200,411.00 | B3*48.7 | 250,300 |',
      '| 4004_Store 40000000000000 | 200,548.00 | B4*48.7 | 250,400 |'
    ].join('\n')
  )

  const tsv = buildHandoff({
    sheetName: sheet.name,
    range: 'A1:D4',
    format: 'tsv',
    cells: selectionToTsv(sheet, box)
  })
  // TSV goes into a fence: pasted loose into a chat, its tabs read as stray whitespace.
  assert.equal(
    tsv,
    [
      'Таблица из XLSX: лист «Ledger», диапазон A1:D4',
      '',
      '```tsv',
      'Store\tSales\tRevenue\tPlan',
      '4002_Store 20000000000000\t200,274.00\tB2*48.7\t250,200',
      '4003_Store 30000000000000\t200,411.00\tB3*48.7\t250,300',
      '4004_Store 40000000000000\t200,548.00\tB4*48.7\t250,400',
      '```'
    ].join('\n')
  )
})

test('the context line names the file only when a caller knows it', () => {
  const body = '| a |\n| --- |\n| 1 |'
  const without = buildHandoff({ sheetName: 'Sheet1', range: 'A1', format: 'markdown', cells: body, fileName: null })
  assert.equal(without, `Таблица из XLSX: лист «Sheet1», диапазон A1\n\n${body}`)

  const withFile = buildHandoff({ sheetName: 'Sheet1', range: 'A1', format: 'markdown', cells: body, fileName: 'books.xlsx' })
  assert.equal(withFile, `Таблица из XLSX: лист «Sheet1», диапазон A1, файл «books.xlsx»\n\n${body}`)
})

test('nothing to send stays nothing: no body, no header', () => {
  assert.equal(buildHandoff({ sheetName: 'Sheet1', range: 'A1', format: 'markdown', cells: '' }), '')
  assert.equal(buildHandoff({ sheetName: 'Sheet1', range: 'A1', format: 'tsv' }), '')
  assert.equal(buildHandoff(), '')
})

test('an unknown format is the Markdown default, not a broken frame', () => {
  const text = buildHandoff({ sheetName: 'S', range: 'A1', format: 'csv', cells: 'x' })
  assert.equal(text, 'Таблица из XLSX: лист «S», диапазон A1\n\nx')
})

test('text under the limit is left exactly as it was', () => {
  const short = 'Таблица из XLSX: лист «S»\n\n| a |'
  assert.deepEqual(truncateText(short), { text: short, truncated: false })
  assert.deepEqual(truncateText(short, short.length), { text: short, truncated: false })
})

test('an over-long handoff is cut on a line boundary and says so', () => {
  const row = (index) => `| row ${index} | ${'x'.repeat(40)} |`
  const body = Array.from({ length: 100 }, (_, index) => row(index)).join('\n')
  assert.ok(body.length > TEXT_LIMIT)

  const cut = truncateText(body)
  assert.equal(cut.truncated, true)
  assert.ok(cut.text.length <= TEXT_LIMIT, `expected <= ${TEXT_LIMIT}, got ${cut.text.length}`)
  assert.ok(cut.text.endsWith(`…обрезано по лимиту ${TEXT_LIMIT} символов`))
  // Every surviving line is a whole table row: no half row is left looking like data.
  const lines = cut.text.split('\n')
  const rows = lines.slice(0, -1)
  assert.ok(rows.length > 0)
  for (const line of rows) assert.equal(line, row(Number(/\d+/.exec(line)[0])))
  assert.equal(lines.at(-1), `…обрезано по лимиту ${TEXT_LIMIT} символов`)
})

test('a first line longer than the limit is cut hard instead of dropped', () => {
  const cut = truncateText('y'.repeat(5000), 100)
  assert.equal(cut.truncated, true)
  assert.equal(cut.text.length, 100)
  assert.ok(cut.text.startsWith('y'))
})

test('no limit means no room for anything but the truncation mark', () => {
  assert.deepEqual(truncateText('abc', 0), { text: '', truncated: true })
  assert.deepEqual(truncateText('', 0), { text: '', truncated: false })
})

test('agents are ranked working, waiting, blocked, then done, newest state first', () => {
  const agents = [
    agent('done-old', 'done', 10),
    agent('blocked', 'blocked', 30),
    agent('working-old', 'working', 20),
    agent('waiting', 'waiting', 40),
    agent('done-new', 'done', 50),
    agent('working-new', 'working', 60)
  ]
  assert.deepEqual(
    rankAgents(agents).map((entry) => entry.terminalId),
    ['working-new', 'working-old', 'waiting', 'blocked', 'done-new', 'done-old']
  )
})

test('agents with an equal state and instant keep the host order, and junk is skipped', () => {
  const agents = [agent('b', 'working', 5), null, agent('a', 'working', 5), 'x', agent('c', 'working', Number.NaN)]
  assert.deepEqual(
    rankAgents(agents).map((entry) => entry.terminalId),
    ['b', 'a', 'c']
  )
  assert.deepEqual(rankAgents(null), [])
})

test('a send request needs a terminal and text, and Enter is opt-in', () => {
  assert.deepEqual(readSend(null), { ok: false, reason: 'agent.send needs an object' })
  assert.equal(readSend({ text: 'x' }).ok, false)
  assert.equal(readSend({ terminalId: '', text: 'x' }).ok, false)
  assert.equal(readSend({ terminalId: 42, text: 'x' }).ok, false)
  assert.equal(readSend({ terminalId: 't1' }).ok, false)
  assert.equal(readSend({ terminalId: 't1', text: '' }).ok, false)
  assert.equal(readSend({ terminalId: 't1', text: 'x', enter: 'yes' }).ok, false)
  assert.deepEqual(readSend({ terminalId: 't1', text: 'x' }), { ok: true, terminalId: 't1', text: 'x', enter: false })
  assert.deepEqual(readSend({ terminalId: 't1', text: 'x', enter: true }), { ok: true, terminalId: 't1', text: 'x', enter: true })
})

test('agent.list asks the host once and hands back the ranked list', async () => {
  const host = fakeOrca({
    'agents.list': {
      agents: [agent('finished', 'done', 1, 'Builder'), agent('busy', 'working', 2, 'Reviewer')]
    }
  })
  activate(host)
  const answer = await host.handlers.get('agent.list')()

  assert.deepEqual(host.calls.map((call) => call.method), ['agents.list'])
  assert.deepEqual(answer.agents.map((entry) => entry.terminalId), ['busy', 'finished'])
})

test('agent.send types the text without pressing Enter unless asked', async () => {
  const host = fakeOrca({ 'terminal.sendText': { accepted: true } })
  activate(host)

  const answer = await host.handlers.get('agent.send')({ terminalId: 'term-1', text: 'Таблица из XLSX' })
  assert.deepEqual(answer, { accepted: true, sent: 'Таблица из XLSX', truncated: false })
  assert.deepEqual(host.calls.at(-1), {
    method: 'terminal.sendText',
    params: { terminalId: 'term-1', text: 'Таблица из XLSX', enter: false }
  })
})

test('agent.send truncates at the host limit before the call and reports it', async () => {
  const host = fakeOrca({ 'terminal.sendText': { accepted: true } })
  activate(host)
  const long = Array.from({ length: 200 }, (_, index) => `| line ${index} | ${'z'.repeat(30)} |`).join('\n')

  const answer = await host.handlers.get('agent.send')({ terminalId: 'term-1', text: long, enter: true })
  const sent = host.calls.at(-1).params.text

  assert.equal(answer.accepted, true)
  assert.equal(answer.truncated, true)
  assert.equal(sent, answer.sent)
  assert.ok(sent.length <= TEXT_LIMIT)
  assert.equal(host.calls.at(-1).params.enter, true)
})

test('agent.send refuses a missing terminal or empty text without calling the host', async () => {
  const host = fakeOrca({ 'terminal.sendText': { accepted: true } })
  activate(host)

  const noTerminal = await host.handlers.get('agent.send')({ terminalId: '', text: 'x' })
  assert.equal(noTerminal.accepted, false)
  assert.equal(noTerminal.sent, '')
  assert.ok(noTerminal.reason)

  const noText = await host.handlers.get('agent.send')({ terminalId: 'term-1', text: '' })
  assert.equal(noText.accepted, false)
  assert.ok(noText.reason)

  assert.deepEqual(host.calls, [])
})

test('a terminal that refuses, or a host that throws, comes back as a refusal', async () => {
  const refused = fakeOrca({ 'terminal.sendText': { accepted: false } })
  activate(refused)
  const answer = await refused.handlers.get('agent.send')({ terminalId: 'term-1', text: 'x' })
  assert.equal(answer.accepted, false)
  assert.equal(answer.reason, 'Терминал не принял текст')

  const broken = fakeOrca({ 'terminal.sendText': new Error('terminal not found') })
  activate(broken)
  const failed = await broken.handlers.get('agent.send')({ terminalId: 'term-1', text: 'x' })
  assert.equal(failed.accepted, false)
  assert.equal(failed.reason, 'terminal not found')
})
