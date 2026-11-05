// The zip slice's own tests. `node --test test/zip.test.mjs`
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test } from 'node:test'
import ExcelJS from 'exceljs'
import { deflateSync } from 'fflate'
import {
  crc32,
  readZip,
  writeZip,
  OFFICE_ZIP_MAX_ENTRIES,
  OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES,
} from '../src/write/zip.js'

const FIXTURES = new URL('./fixtures/', import.meta.url)
const read = (name) => new Uint8Array(readFileSync(new URL(name, FIXTURES)))

// Ground truth from python `zipfile.ZipFile(fixture).infolist()`: name, compressed size,
// uncompressed size, CRC. Sizes here are what the file really holds, not what we compute.
const STYLED = [
  ['docProps/app.xml', 149, 205, 1213056838],
  ['docProps/core.xml', 234, 459, 2577950353],
  ['xl/theme/theme1.xml', 1552, 10140, 597449881],
  ['xl/worksheets/sheet1.xml', 16541, 113514, 3556433766],
  ['xl/styles.xml', 846, 4531, 3298874202],
  ['_rels/.rels', 192, 534, 2330675127],
  ['xl/workbook.xml', 307, 553, 3801309080],
  ['xl/_rels/workbook.xml.rels', 173, 507, 3135499059],
  ['[Content_Types].xml', 283, 983, 2218952347],
]

test('crc32 matches the standard checksums', () => {
  assert.equal(crc32(new Uint8Array(0)), 0)
  assert.equal(crc32(new TextEncoder().encode('The quick brown fox jumps over the lazy dog')), 0x414fa339)
  assert.equal(crc32(new TextEncoder().encode('123456789')), 0xcbf43926)
})

test('readZip reads every entry of styled.xlsx and keeps its compressed bytes', () => {
  const bytes = read('styled.xlsx')
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.equal(zip.entries.length, 9)
  assert.deepEqual(
    zip.entries.map((entry) => entry.name),
    STYLED.map(([name]) => name),
  )
  zip.entries.forEach((entry, index) => {
    const [name, compressedSize, uncompressedSize, crc] = STYLED[index]
    assert.equal(entry.name, name)
    assert.ok(entry.raw instanceof Uint8Array, `${name}: raw is missing`)
    assert.equal(entry.raw.byteLength, compressedSize, `${name}: compressed size`)
    assert.equal(entry.data.byteLength, uncompressedSize, `${name}: uncompressed size`)
    assert.equal(entry.crc, crc, `${name}: crc`)
    assert.equal(entry.method, 8, `${name}: method`)
    assert.equal(crc32(entry.data), crc, `${name}: data matches its crc`)
  })
})

test('writeZip keeps untouched payload bytes identical', () => {
  const original = read('styled.xlsx')
  const zip = readZip(original)
  assert.equal(zip.status, 'ok')
  const originalRaw = zip.entries.map((entry) => entry.raw.slice())

  const rebuilt = writeZip(zip.entries)
  assert.equal(rebuilt.byteLength, original.byteLength, 'a rewrite must not change the file size')

  const again = readZip(rebuilt)
  assert.equal(again.status, 'ok', JSON.stringify(again))
  assert.equal(again.entries.length, 9)
  again.entries.forEach((entry, index) => {
    assert.equal(entry.name, zip.entries[index].name)
    assert.deepEqual(entry.raw, originalRaw[index], `${entry.name}: compressed bytes changed`)
    assert.deepEqual(entry.data, zip.entries[index].data, `${entry.name}: content changed`)
  })

  // Nothing outside a payload may change, and for this fixture the regenerated headers happen to
  // reproduce the original words too, so the whole file comes back byte for byte.
  const ranges = headerRanges(original)
  let differing = 0
  const strayed = []
  for (let index = 0; index < original.byteLength; index++) {
    if (original[index] === rebuilt[index]) continue
    differing++
    if (!ranges.some(([from, to]) => index >= from && index < to)) strayed.push(index)
  }
  assert.deepEqual(strayed, [], 'a payload byte changed on a rewrite without edits')
  assert.equal(differing, 0, `${differing} bytes changed on a rewrite without edits`)
  assert.deepEqual(rebuilt, original)
})

test('macro.xlsm keeps xl/vbaProject.bin byte for byte', () => {
  const original = read('macro.xlsm')
  const zip = readZip(original)
  assert.equal(zip.status, 'ok')
  const macro = zip.entries.find((entry) => entry.name === 'xl/vbaProject.bin')
  assert.ok(macro, 'vbaProject.bin is missing from the fixture')
  assert.equal(macro.data.byteLength, 10240)
  const originalMacroData = macro.data.slice()
  const originalMacroRaw = macro.raw.slice()

  const rebuilt = writeZip(zip.entries)
  const again = readZip(rebuilt)
  assert.equal(again.status, 'ok')
  const after = again.entries.find((entry) => entry.name === 'xl/vbaProject.bin')
  assert.deepEqual(after.data, originalMacroData, 'macro payload changed')
  assert.deepEqual(after.raw, originalMacroRaw, 'macro compressed run changed')
  assert.equal(after.crc, macro.crc)

  // The whole archive survives a rewrite: names, order and content.
  assert.deepEqual(
    again.entries.map((entry) => entry.name),
    zip.entries.map((entry) => entry.name),
  )
  again.entries.forEach((entry, index) => {
    assert.deepEqual(entry.data, zip.entries[index].data, `${entry.name}: content changed`)
  })
})

test('a changed entry is re-encoded while the rest stay byte for byte', () => {
  const original = read('styled.xlsx')
  const zip = readZip(original)
  const target = 'xl/workbook.xml'
  const index = zip.entries.findIndex((entry) => entry.name === target)
  const replacement = new TextEncoder().encode('<workbook>edited</workbook>')
  const entries = zip.entries.map((entry, at) =>
    at === index ? { ...entry, data: replacement, changed: true } : entry,
  )

  const rebuilt = writeZip(entries)
  const again = readZip(rebuilt)
  assert.equal(again.status, 'ok', JSON.stringify(again))
  assert.equal(again.entries.length, zip.entries.length)

  const edited = again.entries[index]
  assert.equal(edited.name, target)
  assert.deepEqual(edited.data, replacement)
  assert.equal(edited.crc, crc32(replacement))
  assert.notEqual(edited.crc, zip.entries[index].crc)
  assert.equal(edited.raw.byteLength, deflateSync(replacement).byteLength)

  again.entries.forEach((entry, at) => {
    if (at === index) return
    assert.deepEqual(entry.raw, zip.entries[at].raw, `${entry.name}: payload was re-compressed`)
    assert.deepEqual(entry.data, zip.entries[at].data)
    assert.equal(entry.crc, zip.entries[at].crc)
  })
})

test('round-trip keeps names and content for entries we build ourselves', () => {
  const encoder = new TextEncoder()
  const entries = [
    { name: '[Content_Types].xml', data: encoder.encode('<Types/>'), method: 0, crc: 0, raw: null, time: 0, date: 0 },
    { name: 'xl/worksheets/лист.xml', data: encoder.encode('<worksheet/>'.repeat(200)), method: 8, crc: 0, raw: null, time: 0, date: 0 },
    { name: 'empty.txt', data: new Uint8Array(0), method: 8, crc: 0, raw: null, time: 0, date: 0 },
  ]
  const bytes = writeZip(entries)
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.deepEqual(
    zip.entries.map((entry) => entry.name),
    entries.map((entry) => entry.name),
  )
  zip.entries.forEach((entry, index) => {
    assert.deepEqual(entry.data, entries[index].data, `${entry.name}: content`)
    assert.equal(entry.crc, crc32(entries[index].data))
    assert.equal(entry.method, entries[index].method)
    assert.ok(entry.raw instanceof Uint8Array)
  })
  // UTF-8 names read back unchanged, and stored entries keep raw === payload.
  assert.equal(zip.entries[0].raw.byteLength, entries[0].data.byteLength)

  // A second rewrite of an already rewritten archive is still stable.
  const twice = readZip(writeZip(zip.entries))
  assert.equal(twice.status, 'ok')
  assert.deepEqual(
    twice.entries.map((entry) => entry.data),
    zip.entries.map((entry) => entry.data),
  )
})

test('writeZip drops the data-descriptor bit and keeps the other flag bits', () => {
  const encoder = new TextEncoder()
  const data = encoder.encode('<a/>')
  const entry = (name, flags) => ({ name, data, method: 0, crc: 0, raw: null, time: 0, date: 0, flags })
  const bytes = writeZip([
    entry('a.xml', 0x0808),
    entry('b.xml', 0),
    entry('c.xml', 0x0002),
    entry('лист.xml', 0),
  ])
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok')
  assert.equal(zip.entries[0].flags, 0x0800, 'bit 3 cleared, UTF-8 bit kept')
  assert.equal(zip.entries[1].flags, 0, 'an ASCII name without the UTF-8 bit keeps it unset')
  assert.equal(zip.entries[2].flags, 0x0002, 'unrelated general-purpose bits survive')
  assert.equal(zip.entries[3].flags, 0x0800, 'a non-ASCII name is flagged as UTF-8')
  assert.equal(zip.entries[3].name, 'лист.xml')
})

test('a declared size that lies is caught by counting the real bytes', () => {
  const bomb = zeroes(4 * 1024 * 1024)
  const bytes = writeZip([{ name: 'bomb.bin', data: bomb, method: 8, crc: 0, raw: null, time: 0, date: 0 }])
  // The header claims 64 bytes; the deflate stream expands to 4 MiB.
  const lying = patchDeclaredSize(bytes, 64)
  const zip = readZip(lying, 1024 * 1024)
  assert.notEqual(zip.status, 'ok')
  assert.equal(zip.status, 'too-large')

  // The same archive read with room enough is refused as corrupt, because the sizes disagree.
  assert.equal(readZip(lying).status, 'corrupt')
})

test('a declared size over the budget is refused before anything is inflated', () => {
  const bytes = writeZip([
    { name: 'big.bin', data: new TextEncoder().encode('short'), method: 8, crc: 0, raw: null, time: 0, date: 0 },
  ])
  const lying = patchDeclaredSize(bytes, 0xfffffff0)
  const refused = readZip(lying)
  assert.equal(refused.status, 'too-large')
  assert.match(refused.reason, /big\.bin/)

  // The same guard holds against the default budget, one byte over it.
  assert.equal(readZip(patchDeclaredSize(bytes, OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES + 1)).status, 'too-large')
  assert.equal(readZip(bytes).status, 'ok')
})

test('more than OFFICE_ZIP_MAX_ENTRIES entries are refused', () => {
  const data = new Uint8Array([0x41])
  const entries = Array.from({ length: OFFICE_ZIP_MAX_ENTRIES }, (unused, index) => ({
    name: `part-${index}.xml`,
    data,
    method: 0,
    crc: 0,
    raw: null,
    time: 0,
    date: 0,
  }))
  // The writer refuses to build an archive over the cap, so an archive exactly at it has the
  // count overwritten: the reader must not believe the end record.
  const atLimit = writeZip(entries)
  const zip = readZip(atLimit)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.equal(zip.entries.length, OFFICE_ZIP_MAX_ENTRIES)

  assert.equal(readZip(patchEntryCount(atLimit, OFFICE_ZIP_MAX_ENTRIES + 1)).status, 'too-large')
})

test('writeZip refuses more entries than the reader accepts', () => {
  const entries = Array.from({ length: OFFICE_ZIP_MAX_ENTRIES + 1 }, (unused, index) => ({
    name: `part-${index}.xml`,
    data: new Uint8Array([0x41]),
    method: 0,
    crc: 0,
    raw: null,
    time: 0,
    date: 0,
  }))
  // Writing past the reader's cap would produce a workbook the editor itself cannot open.
  assert.throws(
    () => writeZip(entries),
    (error) =>
      error instanceof RangeError &&
      error.message.includes(`at most ${OFFICE_ZIP_MAX_ENTRIES} zip entries`),
  )
})

test('an entry that keeps its sizes in a data descriptor still reads', () => {
  const encoder = new TextEncoder()
  const a = encoder.encode('<a/>'.repeat(100))
  const b = encoder.encode('<b/>')
  const bytes = writeZip([
    { name: 'a.xml', data: a, method: 8, crc: 0, raw: null, time: 0, date: 0 },
    { name: 'b.xml', data: b, method: 8, crc: 0, raw: null, time: 0, date: 0 },
  ])
  // Streaming writers zero the local sizes and flag bit 3; the real values live in the directory.
  const view = new DataView(bytes.buffer)
  const flags = view.getUint16(6, true)
  view.setUint16(6, flags | 0x0008, true)
  view.setUint32(14, 0, true)
  view.setUint32(18, 0, true)
  view.setUint32(22, 0, true)

  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.deepEqual(zip.entries[0].data, a)
  assert.deepEqual(zip.entries[1].data, b)
  assert.equal(view.getUint16(6, true) & 0x0008, 0x0008, 'the local header was marked with bit 3')
  assert.equal(view.getUint32(18, true), 0, 'and its compressed size was zeroed')

  const rewritten = readZip(writeZip(zip.entries))
  assert.equal(rewritten.status, 'ok')
  assert.equal(rewritten.entries[0].flags & 0x0008, 0, 'the rewrite states its sizes in the header')
  assert.deepEqual(rewritten.entries[0].data, a)
})

test('reading works on a view that does not start at the buffer origin', () => {
  const original = read('styled.xlsx')
  const padded = new Uint8Array(original.byteLength + 128)
  padded.set(original, 96)
  const view = padded.subarray(96, 96 + original.byteLength)
  const zip = readZip(view)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.equal(zip.entries.length, 9)
  const rebuilt = writeZip(zip.entries)
  assert.deepEqual(rebuilt, original)
})

test('an empty archive is ok and stays empty', () => {
  const bytes = writeZip([])
  assert.equal(bytes.byteLength, 22)
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok')
  assert.deepEqual(zip.entries, [])
  assert.equal(readZip(writeZip(zip.entries)).status, 'ok')
})

test('writeZip refuses an untouched entry whose data no longer matches its CRC', () => {
  const original = read('styled.xlsx')
  const zip = readZip(original)
  const index = zip.entries.findIndex((entry) => entry.name === 'docProps/app.xml')
  const replacement = new TextEncoder().encode('ignored')
  const swapped = zip.entries.map((entry, at) =>
    at === index ? { ...entry, data: replacement } : entry,
  )
  // Without `changed` the writer copies `raw`, but it must not copy a CRC that `data` no longer
  // matches: that would hide the damage behind a fresh, apparently valid header.
  assert.throws(() => writeZip(swapped), /docProps\/app\.xml: CRC mismatch/)

  // Marked as changed, the same data is genuinely re-encoded and the CRC is recomputed.
  const changed = zip.entries.map((entry, at) =>
    at === index ? { ...entry, data: replacement, changed: true } : entry,
  )
  const again = readZip(writeZip(changed))
  assert.equal(again.status, 'ok', JSON.stringify(again))
  assert.deepEqual(again.entries[index].data, replacement)
  assert.equal(again.entries[index].crc, crc32(replacement))
})

test('stored entries keep their raw bytes and survive a rewrite', () => {
  const encoder = new TextEncoder()
  const body = encoder.encode('mimetype body')
  const entries = [
    { name: 'mimetype', data: body, method: 0, crc: crc32(body), raw: null, time: 0, date: 0 },
    { name: 'content.xml', data: encoder.encode('<office/>'.repeat(50)), method: 8, crc: 0, raw: null, time: 0, date: 0 },
  ]
  const first = readZip(writeZip(entries))
  assert.equal(first.status, 'ok', JSON.stringify(first))
  assert.equal(first.entries[0].method, 0, 'stored entries stay stored')
  assert.deepEqual(first.entries[0].raw, body, 'for a stored entry raw is the payload itself')
  assert.deepEqual(first.entries[0].data, body)
  assert.deepEqual(first.entries[1].raw, first.entries[1].raw.slice(), 'deflated entries carry a run')
  assert.notDeepEqual(first.entries[1].raw, first.entries[1].data)

  const second = writeZip(first.entries)
  assert.deepEqual(second, writeZip(entries), 'a rewrite reproduces the archive it read')
})

test('writeZip refuses entries it cannot address without ZIP64', () => {
  const data = new Uint8Array(0)
  const tooMany = Array.from({ length: 0x10000 }, (unused, index) => ({
    name: `f${index}`,
    data,
    method: 0,
    crc: 0,
    raw: null,
    time: 0,
    date: 0,
  }))
  assert.throws(() => writeZip(tooMany), RangeError)

  const longName = 'x'.repeat(0x10000)
  assert.throws(
    () => writeZip([{ name: longName, data, method: 0, crc: 0, raw: null, time: 0, date: 0 }]),
    RangeError,
  )
  assert.throws(() => writeZip([{ name: 'no-data', method: 0, crc: 0, raw: null }]), TypeError)
})

test('ZIP64 markers and unknown compression methods are refused, not guessed', () => {
  const encoder = new TextEncoder()
  const bytes = writeZip([
    { name: 'a.xml', data: encoder.encode('<a/>'), method: 8, crc: 0, raw: null, time: 0, date: 0 },
  ])
  const withMarker = bytes.slice()
  new DataView(withMarker.buffer).setUint32(findCentral(withMarker) + 20, 0xffffffff, true)
  assert.equal(readZip(withMarker).status, 'too-large')

  const unsupported = bytes.slice()
  new DataView(unsupported.buffer).setUint16(findCentral(unsupported) + 10, 9, true)
  const refused = readZip(unsupported)
  assert.equal(refused.status, 'corrupt')
  assert.match(refused.reason, /method 9/)
})

test('garbage bytes are reported as corrupt instead of throwing', () => {
  assert.deepEqual(
    { status: readZip(new Uint8Array(0)).status },
    { status: 'corrupt' },
  )
  assert.equal(readZip(new TextEncoder().encode('this is not a zip at all')).status, 'corrupt')
  const noise = new Uint8Array(4096)
  for (let index = 0; index < noise.length; index++) noise[index] = (index * 37 + 11) & 0xff
  assert.equal(readZip(noise).status, 'corrupt')

  // A real archive cut in half loses its central directory.
  const original = read('styled.xlsx')
  assert.equal(readZip(original.subarray(0, original.byteLength - 200)).status, 'corrupt')
  // A damaged deflate stream inside an intact directory.
  const zip = readZip(original)
  const entry = zip.entries[2]
  assert.equal(entry.name, 'xl/theme/theme1.xml')
  const damaged = original.slice()
  const at = entry.raw.byteOffset - original.byteOffset
  for (let index = 0; index < 40; index++) damaged[at + index] = 0xff
  assert.equal(readZip(damaged).status, 'corrupt')

  // Every refusal carries a readable reason and no entries.
  const refused = readZip(new Uint8Array([1, 2, 3]))
  assert.equal(refused.status, 'corrupt')
  assert.equal(typeof refused.reason, 'string')
  assert.ok(refused.reason.length > 0)
})

test('readZip rejects a stored payload whose CRC does not match its header', () => {
  const body = new TextEncoder().encode('checksum me')
  const bytes = writeZip([
    { name: 'stored.txt', data: body, method: 0, crc: 0, raw: null, time: 0, date: 0 },
  ])
  // A stored payload starts right after the 30-byte local header and the entry name; inverting a
  // byte leaves every size intact, so only the checksum can notice.
  const damaged = bytes.slice()
  damaged[30 + 'stored.txt'.length] ^= 0x01
  const result = readZip(damaged)
  assert.equal(result.status, 'corrupt')
  assert.equal(result.reason, 'stored.txt: CRC mismatch')
})

test('readZip rejects a deflated payload whose CRC does not match its header', () => {
  const body = new TextEncoder().encode('<worksheet>'.repeat(64))
  const bytes = writeZip([
    { name: 'sheet.xml', data: body, method: 8, crc: 0, raw: null, time: 0, date: 0 },
  ])
  // The stream is intact and inflates to the declared size; only the claimed checksum is wrong.
  const lying = bytes.slice()
  new DataView(lying.buffer).setUint32(findCentral(lying) + 16, crc32(body) ^ 0xffffffff, true)
  const result = readZip(lying)
  assert.equal(result.status, 'corrupt')
  assert.equal(result.reason, 'sheet.xml: CRC mismatch')

  // Inverting a byte inside the deflate stream is never reported as ok, whatever it does to it.
  const inverted = bytes.slice()
  const view = new DataView(inverted.buffer)
  const localOffset = view.getUint32(findCentral(inverted) + 42, true)
  const dataStart =
    localOffset + 30 + view.getUint16(localOffset + 26, true) + view.getUint16(localOffset + 28, true)
  inverted[dataStart + 1] ^= 0xff
  assert.equal(readZip(inverted).status, 'corrupt')
})

test('raw is a copy, not a window into the caller buffer', () => {
  const original = read('styled.xlsx')
  const zip = readZip(original)
  assert.equal(zip.status, 'ok')
  const expected = writeZip(zip.entries)
  const raws = zip.entries.map((entry) => entry.raw.slice())
  for (const entry of zip.entries) assert.notEqual(entry.raw.buffer, original.buffer)

  // Rewriting the bytes that were read must not change what a save produces.
  original.fill(0xff)
  assert.deepEqual(writeZip(zip.entries), expected)
  zip.entries.forEach((entry, index) => {
    assert.deepEqual(entry.raw, raws[index], `${entry.name}: raw changed with the input buffer`)
  })

  // A Buffer is a Uint8Array too, but its own `slice` returns a view; raw must still be a copy.
  const buffer = Buffer.from(read('styled.xlsx'))
  const fromBuffer = readZip(buffer)
  assert.equal(fromBuffer.status, 'ok')
  const bufferRaws = fromBuffer.entries.map((entry) => entry.raw.slice())
  buffer.fill(0xff)
  fromBuffer.entries.forEach((entry, index) => {
    assert.deepEqual(entry.raw, bufferRaws[index], `${entry.name}: raw followed a Buffer input`)
  })
})

test('the entry count in the end record is checked against the directory', () => {
  const encoder = new TextEncoder()
  const bytes = writeZip(
    ['a', 'b', 'c'].map((name) => ({
      name: `${name}.txt`,
      data: encoder.encode(name),
      method: 0,
      crc: 0,
      raw: null,
      time: 0,
      date: 0,
    })),
  )
  assert.equal(readZip(bytes).status, 'ok')

  // A count that is too small would silently drop records; walking the directory must land on its
  // declared end instead of trusting the number.
  for (const count of [0, 1, 2]) {
    const result = readZip(patchEntryCount(bytes, count))
    assert.equal(result.status, 'corrupt', `a count of ${count} was believed`)
    assert.match(result.reason, /central directory/)
  }
  // A count that is too large cannot read into the end record either.
  assert.equal(readZip(patchEntryCount(bytes, 4)).status, 'corrupt')
})

test('a name without the UTF-8 flag is decoded as CP437', () => {
  const encoder = new TextEncoder()
  const bytes = buildStoredZip([
    {
      nameBytes: concatBytes(encoder.encode('caf'), Uint8Array.of(0x82), encoder.encode('.txt')),
      data: encoder.encode('acute'),
    },
    {
      nameBytes: concatBytes(encoder.encode('caf'), Uint8Array.of(0x8a), encoder.encode('.txt')),
      data: encoder.encode('grave'),
    },
  ])
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.deepEqual(zip.entries.map((entry) => entry.name), ['café.txt', 'cafè.txt'])

  // The rewrite stores both names as UTF-8; they stay distinct and keep their own payloads.
  const again = readZip(writeZip(zip.entries))
  assert.equal(again.status, 'ok', JSON.stringify(again))
  assert.deepEqual(again.entries.map((entry) => entry.name), ['café.txt', 'cafè.txt'])
  assert.deepEqual(
    again.entries.map((entry) => entry.data),
    zip.entries.map((entry) => entry.data),
  )
})

test('an Info-ZIP Unicode Path extra field supplies the name', () => {
  const encoder = new TextEncoder()
  const nameBytes = encoder.encode('naive.txt')
  const bytes = buildStoredZip([
    { nameBytes, data: encoder.encode('body'), extra: unicodePathExtra(nameBytes, 'naïve.txt') },
  ])
  const zip = readZip(bytes)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  assert.equal(zip.entries[0].name, 'naïve.txt')
})

test('an entry with an empty name is corrupt instead of failing later', () => {
  const bytes = buildStoredZip([{ nameBytes: new Uint8Array(0), data: new Uint8Array(0) }])
  const result = readZip(bytes)
  assert.equal(result.status, 'corrupt')
  assert.match(result.reason, /empty name/)
})

test('an ExcelJS workbook round-trips byte for byte, version needed included', async () => {
  const workbook = new ExcelJS.Workbook()
  const sheet = workbook.addWorksheet('Sheet1')
  sheet.getCell('A1').value = 'version'
  sheet.getCell('B2').value = 42
  const original = new Uint8Array(await workbook.xlsx.writeBuffer())
  const zip = readZip(original)
  assert.equal(zip.status, 'ok', JSON.stringify(zip))
  // ExcelJS asks for version 1.0; the reader carries it through instead of stamping 2.0 over it.
  assert.ok(zip.entries.length > 0)
  assert.ok(
    zip.entries.every((entry) => entry.versionNeeded === 10),
    'ExcelJS writes version needed 1.0',
  )

  const rebuilt = writeZip(zip.entries)
  assert.equal(rebuilt.byteLength, original.byteLength)
  const differing = []
  for (let index = 0; index < original.byteLength; index++) {
    if (original[index] !== rebuilt[index]) differing.push(index)
  }
  assert.deepEqual(differing, [], `${differing.length} bytes changed on a rewrite without edits`)

  // A record the caller invents has no original value, so the writer stamps version 2.0.
  const fresh = readZip(
    writeZip([{ name: 'new.txt', data: new Uint8Array([1]), method: 0, crc: 0, raw: null, time: 0, date: 0 }]),
  )
  assert.equal(fresh.status, 'ok')
  assert.equal(fresh.entries[0].versionNeeded, 20)
})

function zeroes(size) {
  return new Uint8Array(size)
}

function concatBytes(...arrays) {
  const out = new Uint8Array(arrays.reduce((sum, array) => sum + array.length, 0))
  let at = 0
  for (const array of arrays) {
    out.set(array, at)
    at += array.length
  }
  return out
}

/** Info-ZIP Unicode Path extra field (0x7075) for `nameBytes`, carrying `unicodeName`. */
function unicodePathExtra(nameBytes, unicodeName) {
  const name = new TextEncoder().encode(unicodeName)
  const extra = new Uint8Array(4 + 1 + 4 + name.length)
  const view = new DataView(extra.buffer)
  view.setUint16(0, 0x7075, true)
  view.setUint16(2, 1 + 4 + name.length, true)
  extra[4] = 1
  view.setUint32(5, crc32(nameBytes), true)
  extra.set(name, 9)
  return extra
}

/**
 * A zip assembled by hand, for records the writer cannot emit: names in another encoding, an empty
 * name, a central-directory extra field. Every entry is stored and has no data descriptor.
 */
function buildStoredZip(records) {
  const locals = []
  const centrals = []
  let localOffset = 0
  for (const record of records) {
    const nameBytes = record.nameBytes
    const data = record.data ?? new Uint8Array(0)
    const extra = record.extra ?? new Uint8Array(0)
    const checksum = crc32(data)

    const local = new Uint8Array(30 + nameBytes.length + data.length)
    const localView = new DataView(local.buffer)
    localView.setUint32(0, 0x04034b50, true)
    localView.setUint16(4, 20, true)
    localView.setUint16(8, 0, true)
    localView.setUint16(12, 0x21, true)
    localView.setUint32(14, checksum, true)
    localView.setUint32(18, data.length, true)
    localView.setUint32(22, data.length, true)
    localView.setUint16(26, nameBytes.length, true)
    local.set(nameBytes, 30)
    local.set(data, 30 + nameBytes.length)

    const central = new Uint8Array(46 + nameBytes.length + extra.length)
    const centralView = new DataView(central.buffer)
    centralView.setUint32(0, 0x02014b50, true)
    centralView.setUint16(4, 20, true)
    centralView.setUint16(6, 20, true)
    centralView.setUint16(14, 0x21, true)
    centralView.setUint32(16, checksum, true)
    centralView.setUint32(20, data.length, true)
    centralView.setUint32(24, data.length, true)
    centralView.setUint16(28, nameBytes.length, true)
    centralView.setUint16(30, extra.length, true)
    centralView.setUint32(42, localOffset, true)
    central.set(nameBytes, 46)
    central.set(extra, 46 + nameBytes.length)

    locals.push(local)
    centrals.push(central)
    localOffset += local.length
  }

  const centralSize = centrals.reduce((sum, part) => sum + part.length, 0)
  const out = new Uint8Array(localOffset + centralSize + 22)
  let at = 0
  for (const part of [...locals, ...centrals]) {
    out.set(part, at)
    at += part.length
  }
  const view = new DataView(out.buffer)
  view.setUint32(at, 0x06054b50, true)
  view.setUint16(at + 8, records.length, true)
  view.setUint16(at + 10, records.length, true)
  view.setUint32(at + 12, centralSize, true)
  view.setUint32(at + 16, localOffset, true)
  return out
}

/** Offset of the end of central directory record. */
function findEndRecord(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let at = view.byteLength - 22; at >= Math.max(0, view.byteLength - 22 - 0xffff); at--) {
    if (view.getUint32(at, true) === 0x06054b50) return at
  }
  throw new Error('end of central directory record not found')
}

/** Overwrites both entry counts in the end of central directory record. */
function patchEntryCount(bytes, count) {
  const copied = bytes.slice()
  const view = new DataView(copied.buffer, copied.byteOffset, copied.byteLength)
  const end = findEndRecord(copied)
  view.setUint16(end + 8, count, true)
  view.setUint16(end + 10, count, true)
  return copied
}

/** Offset of the single central directory record in a one-entry archive. */
function findCentral(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let at = 0; at + 4 <= view.byteLength; at++) {
    if (view.getUint32(at, true) === 0x02014b50) return at
  }
  throw new Error('central directory not found')
}

/**
 * Overwrites the uncompressed size a single-entry archive declares, in both the local header and
 * the central directory, leaving everything else intact.
 */
function patchDeclaredSize(bytes, size) {
  const copied = bytes.slice()
  const view = new DataView(copied.buffer, copied.byteOffset, copied.byteLength)
  view.setUint32(22, size, true)
  view.setUint32(findCentral(copied) + 24, size, true)
  return copied
}

/**
 * Byte ranges of everything that is not payload: local headers, the central directory and the end
 * record. Written here rather than imported so the test does not trust the code it checks.
 */
function headerRanges(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let end = -1
  for (let at = view.byteLength - 22; at >= Math.max(0, view.byteLength - 22 - 0xffff); at--) {
    if (view.getUint32(at, true) === 0x06054b50) {
      end = at
      break
    }
  }
  assert.ok(end >= 0, 'fixture has no end of central directory record')
  const count = view.getUint16(end + 10, true)
  const centralSize = view.getUint32(end + 12, true)
  const centralOffset = view.getUint32(end + 16, true)
  const ranges = [[centralOffset, centralOffset + centralSize], [end, end + 22]]
  let at = centralOffset
  for (let index = 0; index < count; index++) {
    const nameLength = view.getUint16(at + 28, true)
    const extraLength = view.getUint16(at + 30, true)
    const commentLength = view.getUint16(at + 32, true)
    const localOffset = view.getUint32(at + 42, true)
    const dataStart =
      localOffset +
      30 +
      view.getUint16(localOffset + 26, true) +
      view.getUint16(localOffset + 28, true)
    ranges.push([localOffset, dataStart])
    at += 46 + nameLength + extraLength + commentLength
  }
  return ranges
}
