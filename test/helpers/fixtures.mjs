// Access to the workbook fixtures from a test.
//
// The books are built by `test/helpers/build-fixtures.mjs` and are not committed — a plugin
// repository the catalog accepts carries no spreadsheet binaries. The first call in a test process
// writes the three files into `test/fixtures/`, so `git clone && npm ci && npm test` needs no
// preparation and no Python, and `rm -rf test/fixtures` is a supported starting point.
//
// `node --test` runs each file in its own process, and those processes may overlap: a book is
// written under a temporary name and renamed into place, so a sibling never reads a half-written
// archive.
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import path from 'node:path'

import { buildFixtures, OUTPUT_DIR } from './build-fixtures.mjs'

export const FIXTURES_DIR = OUTPUT_DIR

const NAMES = ['rich.xlsx', 'macro.xlsm', 'styled.xlsx']
let prepared = false

/** Writes every fixture that is missing or out of date; idempotent and safe to race. */
function prepare() {
  if (prepared) return
  mkdirSync(FIXTURES_DIR, { recursive: true })
  for (const [name, bytes] of Object.entries(buildFixtures())) {
    const target = path.join(FIXTURES_DIR, name)
    if (existsSync(target) && Buffer.from(bytes).equals(readFileSync(target))) continue
    const temporary = `${target}.${process.pid}.tmp`
    writeFileSync(temporary, bytes)
    renameSync(temporary, target)
  }
  prepared = true
}

/**
 * Absolute path of one fixture, building the books on first use.
 *
 * @param {'rich.xlsx' | 'macro.xlsm' | 'styled.xlsx'} name
 * @returns {string}
 */
export function fixturePath(name) {
  if (!NAMES.includes(name)) throw new Error(`unknown fixture ${name}`)
  prepare()
  return path.join(FIXTURES_DIR, name)
}

/**
 * Bytes of one fixture — what most tests want, since the reader takes a `Uint8Array`.
 *
 * @param {'rich.xlsx' | 'macro.xlsm' | 'styled.xlsx'} name
 * @returns {Uint8Array}
 */
export function fixtureBytes(name) {
  return new Uint8Array(readFileSync(fixturePath(name)))
}
