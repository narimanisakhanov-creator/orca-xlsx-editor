// The fixtures themselves: what the generator must keep producing, and that it does so identically
// on every run. The specs that read the books assert on their content; this file asserts on the
// generator, so a change to it that silently drops a part or stamps the current time fails here
// instead of in a reader test with a confusing message.
import assert from 'node:assert/strict'
import path from 'node:path'
import { test } from 'node:test'

import { readZip } from '../src/write/zip.js'
import { buildFixtures } from './helpers/build-fixtures.mjs'
import { FIXTURES_DIR, fixtureBytes, fixturePath } from './helpers/fixtures.mjs'

const NAMES = ['rich.xlsx', 'macro.xlsm', 'styled.xlsx']

/** Entry names of one book, through the plugin's own reader. */
function partNames(bytes) {
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok', `unreadable fixture: ${zip.status}`)
  return zip.entries.map((entry) => entry.name)
}

test('the generator builds three books and nothing else', () => {
  assert.deepEqual(Object.keys(buildFixtures()).sort(), [...NAMES].sort())
})

test('every book is byte-identical across two builds', () => {
  const first = buildFixtures()
  const second = buildFixtures()
  for (const name of NAMES) {
    // A zip record carries a DOS timestamp and a deflate stream depends on content alone, so two
    // builds must agree to the byte. A `Date.now()` anywhere in the generator breaks this test.
    assert.ok(first[name].byteLength > 0, `${name}: empty archive`)
    assert.deepEqual(first[name], second[name], `${name} differs between two builds`)
  }
})

test('the three books carry the parts the specs address by name', () => {
  const books = buildFixtures()

  const rich = partNames(books['rich.xlsx'])
  for (const name of [
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/sheet2.xml',
    'xl/worksheets/sheet3.xml',
    '[Content_Types].xml',
  ]) {
    assert.ok(rich.includes(name), `rich.xlsx is missing ${name}`)
  }

  // The macro twin is rich.xlsx plus the project part, so a save can prove it copies it untouched.
  const macro = partNames(books['macro.xlsm'])
  assert.deepEqual(macro, [...rich, 'xl/vbaProject.bin'], 'macro.xlsm is rich.xlsx plus the project')

  const styled = partNames(books['styled.xlsx'])
  assert.deepEqual(styled, [
    'docProps/app.xml',
    'docProps/core.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet1.xml',
    'xl/styles.xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    '[Content_Types].xml',
  ])
})

test('the macro part is a real part with its own content type and a fixed payload', () => {
  const book = buildFixtures()['macro.xlsm']
  const zip = readZip(book)
  assert.equal(zip.status, 'ok')
  const macro = zip.entries.find((entry) => entry.name === 'xl/vbaProject.bin')
  assert.ok(macro, 'vbaProject.bin is missing')
  assert.equal(macro.data.byteLength, 10240)
  // Deterministic stand-in: `index % 256`, so the payload is reproducible and the tests can compare
  // what a save produced against what the file held without storing the book.
  assert.equal(macro.data[0], 0)
  assert.equal(macro.data[255], 255)
  assert.equal(macro.data[256], 0)

  const types = zip.entries.find((entry) => entry.name === '[Content_Types].xml')
  assert.match(
    new TextDecoder().decode(types.data),
    /<Override PartName="\/xl\/vbaProject\.bin" ContentType="application\/vnd\.ms-office\.vbaProject"\/>/,
  )
})

test('the helper writes the books where the specs look for them', () => {
  for (const name of NAMES) {
    assert.equal(fixturePath(name), path.join(FIXTURES_DIR, name))
  }
  // The book on disk is the one the generator builds, not a stale copy from an older run.
  assert.deepEqual(fixtureBytes('styled.xlsx'), buildFixtures()['styled.xlsx'])
  assert.throws(() => fixturePath('nope.xlsx'), /unknown fixture/)
})
