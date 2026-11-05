// The zip a workbook lives in. Reading keeps each entry's exact compressed run (`raw`) so a save
// copies untouched parts byte for byte; only entries marked `changed` are compressed again.
//
// fflate's unprefixed codecs are raw DEFLATE (no zlib/gzip wrapper) and are the ones available in
// the browser build the editor frame is bundled from — there is no `inflateRawSync` there.
import { deflateSync, Inflate } from 'fflate'
import { OFFICE_ZIP_MAX_ENTRIES, OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES } from './zip-budget.js'

export { OFFICE_ZIP_MAX_ENTRIES, OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES }

const LOCAL_HEADER_SIGNATURE = 0x04034b50
const CENTRAL_HEADER_SIGNATURE = 0x02014b50
const END_RECORD_SIGNATURE = 0x06054b50
const LOCAL_HEADER_BYTES = 30
const CENTRAL_HEADER_BYTES = 46
const END_RECORD_BYTES = 22
const MAX_COMMENT_BYTES = 0xffff
const ZIP64_MARKER = 0xffffffff
const ZIP64_ENTRY_MARKER = 0xffff
const STORED = 0
const DEFLATED = 8
const VERSION_20 = 20
// Bit 3 says the sizes follow the payload in a data descriptor. We always write them into the
// header, so that bit is cleared; every other general-purpose bit is carried over.
const DATA_DESCRIPTOR_FLAG = 0x0008
const UTF8_FLAG = 0x0800
// Info-ZIP "Unicode Path" extra field: the real name when the header field is in another encoding.
const EXTRA_UNICODE_PATH = 0x7075
// One compressed chunk may expand ~1000:1, so the byte cap is applied after every inflate step
// instead of trusting the size the archive declares. 4 KiB in keeps peak memory near 4 MiB out.
const INFLATE_CHUNK_BYTES = 1 << 12
const EMPTY = new Uint8Array(0)

const CRC_TABLE = new Uint32Array(256)
for (let n = 0; n < 256; n++) {
  let c = n
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  CRC_TABLE[n] = c >>> 0
}

// Code page 437, bytes 0x80..0xff. A zip name without the UTF-8 flag is defined to use it, and
// refusing accented names — or reading them as mojibake — would rename parts of a real workbook.
const CP437_HIGH = [
  0x00c7, 0x00fc, 0x00e9, 0x00e2, 0x00e4, 0x00e0, 0x00e5, 0x00e7, 0x00ea, 0x00eb, 0x00e8, 0x00ef,
  0x00ee, 0x00ec, 0x00c4, 0x00c5, 0x00c9, 0x00e6, 0x00c6, 0x00f4, 0x00f6, 0x00f2, 0x00fb, 0x00f9,
  0x00ff, 0x00d6, 0x00dc, 0x00a2, 0x00a3, 0x00a5, 0x20a7, 0x0192, 0x00e1, 0x00ed, 0x00f3, 0x00fa,
  0x00f1, 0x00d1, 0x00aa, 0x00ba, 0x00bf, 0x2310, 0x00ac, 0x00bd, 0x00bc, 0x00a1, 0x00ab, 0x00bb,
  0x2591, 0x2592, 0x2593, 0x2502, 0x2524, 0x2561, 0x2562, 0x2556, 0x2555, 0x2563, 0x2551, 0x2557,
  0x255d, 0x255c, 0x255b, 0x2510, 0x2514, 0x2534, 0x252c, 0x251c, 0x2500, 0x253c, 0x255e, 0x255f,
  0x255a, 0x2554, 0x2569, 0x2566, 0x2560, 0x2550, 0x256c, 0x2567, 0x2568, 0x2564, 0x2565, 0x2559,
  0x2558, 0x2552, 0x2553, 0x256b, 0x256a, 0x2518, 0x250c, 0x2588, 0x2584, 0x258c, 0x2590, 0x2580,
  0x03b1, 0x00df, 0x0393, 0x03c0, 0x03a3, 0x03c3, 0x00b5, 0x03c4, 0x03a6, 0x0398, 0x03a9, 0x03b4,
  0x221e, 0x03c6, 0x03b5, 0x2229, 0x2261, 0x00b1, 0x2265, 0x2264, 0x2320, 0x2321, 0x00f7, 0x2248,
  0x00b0, 0x2219, 0x00b7, 0x221a, 0x207f, 0x00b2, 0x25a0, 0x00a0,
]
const UTF8_DECODER = new TextDecoder()

/** CRC-32 of `bytes`, the checksum every zip entry header carries. */
export function crc32(bytes) {
  let crc = 0xffffffff
  for (let index = 0; index < bytes.length; index++) crc = CRC_TABLE[(crc ^ bytes[index]) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

/**
 * @typedef {{ name: string, data: Uint8Array, method: number, crc: number,
 *   raw: Uint8Array | null, time: number, date: number, flags: number,
 *   changed?: boolean, uncompressedSize?: number, compressedSize?: number,
 *   versionMadeBy?: number, versionNeeded?: number, internalAttributes?: number,
 *   externalAttributes?: number }} ZipEntry
 *
 * `raw` is a copy of the exact compressed run found in the file, so a save can copy it verbatim
 * and a caller rewriting its input bytes cannot change an entry behind its back; it is null only
 * for entries the caller invented. `uncompressedSize`, `compressedSize`, `versionMadeBy`,
 * `versionNeeded`, `internalAttributes` and `externalAttributes` are optional metadata the archive
 * arrived with, kept so that rewriting it reproduces bytes the file already had. Callers may ignore
 * all of them.
 */

/**
 * Reads the entries of a zip, keeping each compressed run so it can be written back untouched.
 * Declared sizes are treated as hints: the inflated bytes are counted against `maxBytes`.
 *
 * @returns {{ status: 'ok', entries: ZipEntry[] }
 *   | { status: 'too-large' | 'corrupt', reason: string }}
 */
export function readZip(bytes, maxBytes = OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES) {
  if (!(bytes instanceof Uint8Array)) return refuse('corrupt', 'input is not a byte array')
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = findEndRecord(view)
  if (end < 0) return refuse('corrupt', 'no end of central directory record')

  if (view.getUint16(end + 4, true) !== 0 || view.getUint16(end + 6, true) !== 0) {
    return refuse('corrupt', 'multi-disk archives are not supported')
  }
  const count = view.getUint16(end + 10, true)
  if (view.getUint16(end + 8, true) === ZIP64_ENTRY_MARKER || count === ZIP64_ENTRY_MARKER) {
    return refuse('too-large', 'ZIP64 archives are not supported')
  }
  if (count > OFFICE_ZIP_MAX_ENTRIES) {
    return refuse('too-large', `the workbook has ${count} entries, more than ${OFFICE_ZIP_MAX_ENTRIES}`)
  }
  const centralSize = view.getUint32(end + 12, true)
  const centralOffset = view.getUint32(end + 16, true)
  if (centralSize === ZIP64_MARKER || centralOffset === ZIP64_MARKER) {
    return refuse('too-large', 'ZIP64 archives are not supported')
  }
  if (centralOffset + centralSize > view.byteLength) {
    return refuse('corrupt', 'the central directory runs past the end of the file')
  }

  const limit = Number.isFinite(maxBytes) ? Number(maxBytes) : OFFICE_ZIP_MAX_UNCOMPRESSED_BYTES
  let declaredTotal = 0
  let actualTotal = 0
  const entries = []
  let offset = centralOffset

  for (let index = 0; index < count; index++) {
    if (offset + CENTRAL_HEADER_BYTES > view.byteLength) return refuse('corrupt', 'truncated central directory')
    if (view.getUint32(offset, true) !== CENTRAL_HEADER_SIGNATURE) {
      return refuse('corrupt', 'bad central directory entry signature')
    }
    const versionMadeBy = view.getUint16(offset + 4, true)
    const versionNeeded = view.getUint16(offset + 6, true)
    const flags = view.getUint16(offset + 8, true)
    const method = view.getUint16(offset + 10, true)
    const time = view.getUint16(offset + 12, true)
    const date = view.getUint16(offset + 14, true)
    const crc = view.getUint32(offset + 16, true)
    const compressedSize = view.getUint32(offset + 20, true)
    const declaredSize = view.getUint32(offset + 24, true)
    const nameLength = view.getUint16(offset + 28, true)
    const extraLength = view.getUint16(offset + 30, true)
    const commentLength = view.getUint16(offset + 32, true)
    const internalAttributes = view.getUint16(offset + 36, true)
    const externalAttributes = view.getUint32(offset + 38, true)
    const localOffset = view.getUint32(offset + 42, true)
    // Workbooks this size never need ZIP64, and its real sizes hide in an extra field we do not read.
    if (
      compressedSize === ZIP64_MARKER ||
      declaredSize === ZIP64_MARKER ||
      localOffset === ZIP64_MARKER
    ) {
      return refuse('too-large', 'ZIP64 entries are not supported')
    }
    const nameStart = offset + CENTRAL_HEADER_BYTES
    const extraStart = nameStart + nameLength
    if (extraStart + extraLength + commentLength > view.byteLength) {
      return refuse('corrupt', 'entry metadata runs past the file')
    }
    const nameBytes = bytes.subarray(nameStart, extraStart)
    const name = decodeEntryName(nameBytes, flags, bytes.subarray(extraStart, extraStart + extraLength))
    if (name.length === 0) return refuse('corrupt', 'an entry has an empty name')
    offset = extraStart + extraLength + commentLength

    if (declaredSize > limit - declaredTotal) {
      return refuse('too-large', `${name} declares more than ${limit} bytes of content`)
    }
    declaredTotal += declaredSize

    if (localOffset + LOCAL_HEADER_BYTES > view.byteLength) return refuse('corrupt', `${name}: missing local header`)
    if (view.getUint32(localOffset, true) !== LOCAL_HEADER_SIGNATURE) {
      return refuse('corrupt', `${name}: bad local header signature`)
    }
    const dataStart =
      localOffset + LOCAL_HEADER_BYTES + view.getUint16(localOffset + 26, true) + view.getUint16(localOffset + 28, true)
    if (dataStart + compressedSize > view.byteLength) return refuse('corrupt', `${name}: payload runs past the file`)
    // A copy, not a view: the caller may rewrite its input buffer after reading, and an entry
    // must keep the bytes it was read from. `set` is used instead of `slice` because a Buffer
    // overrides `slice` to return another view into the same memory.
    const raw = new Uint8Array(compressedSize)
    raw.set(bytes.subarray(dataStart, dataStart + compressedSize))

    let data
    if (method === STORED) {
      if (compressedSize !== declaredSize) return refuse('corrupt', `${name}: stored entry with mismatched sizes`)
      // Copied as well so a later in-place edit of `data` cannot silently desynchronise `raw`.
      data = raw.slice()
    } else if (method === DEFLATED) {
      const inflated = inflateWithin(raw, limit - actualTotal)
      if (inflated.status !== 'ok') return refuse(inflated.status, `${name}: ${inflated.reason}`)
      if (inflated.data.byteLength !== declaredSize) {
        return refuse('corrupt', `${name}: inflated to ${inflated.data.byteLength} bytes, header says ${declaredSize}`)
      }
      data = inflated.data
    } else {
      return refuse('corrupt', `${name}: compression method ${method} is not supported`)
    }
    // The header checksum is the only thing that catches a payload damaged in a way the deflate
    // stream itself survives; a workbook with wrong bytes must not open as if it were fine.
    if (crc32(data) !== crc) return refuse('corrupt', `${name}: CRC mismatch`)
    actualTotal += data.byteLength
    entries.push({
      name,
      data,
      method,
      crc,
      raw,
      time,
      date,
      flags,
      // Sizes as the file states them, so a rewrite of an untouched entry does not depend on `data`.
      uncompressedSize: declaredSize,
      compressedSize,
      versionMadeBy,
      versionNeeded,
      internalAttributes,
      externalAttributes,
    })
  }

  // The count in the end record is not trusted on its own: an archive whose directory holds more
  // records than it admits would silently lose the extra ones. Walking the directory must land
  // exactly on its declared end.
  if (entries.length !== count || offset !== centralOffset + centralSize) {
    return refuse('corrupt', 'the central directory does not match the entry count in its end record')
  }

  return { status: 'ok', entries }
}

/**
 * Writes a zip of `entries`, in order. An entry whose `raw` compressed bytes are still there and
 * that is not marked `changed` is copied verbatim together with its recorded CRC and sizes, after
 * checking that `data` still hashes to that CRC; anything else is compressed from `data`.
 */
export function writeZip(entries) {
  if (!Array.isArray(entries)) throw new TypeError('writeZip needs an array of entries')
  // A file this writer emits must be one the editor can open again: the reader caps entry count
  // below the 65535 a 16-bit field could hold, so refuse before building it.
  if (entries.length > OFFICE_ZIP_MAX_ENTRIES) {
    throw new RangeError(
      `a workbook holds at most ${OFFICE_ZIP_MAX_ENTRIES} zip entries, not ${entries.length}`,
    )
  }
  const encoder = new TextEncoder()
  const prepared = entries.map((entry) => prepare(entry, encoder))

  let localBytes = 0
  let centralBytes = 0
  for (const item of prepared) {
    if (item.nameBytes.byteLength > 0xffff) throw new RangeError(`${item.nameBytes.byteLength} byte entry name`)
    if (item.payload.byteLength >= ZIP64_MARKER || item.size >= ZIP64_MARKER) {
      throw new RangeError('an entry of 4 GiB or more needs ZIP64, which this writer does not emit')
    }
    localBytes += LOCAL_HEADER_BYTES + item.nameBytes.byteLength + item.payload.byteLength
    centralBytes += CENTRAL_HEADER_BYTES + item.nameBytes.byteLength
  }
  const centralOffset = localBytes
  if (centralOffset + centralBytes >= ZIP64_MARKER) {
    throw new RangeError('workbook exceeds the 4 GiB a zip without ZIP64 can address')
  }

  const out = new Uint8Array(localBytes + centralBytes + END_RECORD_BYTES)
  const view = new DataView(out.buffer, out.byteOffset, out.byteLength)
  let localAt = 0
  let centralAt = centralOffset

  for (const item of prepared) {
    view.setUint32(localAt, LOCAL_HEADER_SIGNATURE, true)
    view.setUint16(localAt + 4, item.versionNeeded, true)
    view.setUint16(localAt + 6, item.flags, true)
    view.setUint16(localAt + 8, item.method, true)
    view.setUint16(localAt + 10, item.time, true)
    view.setUint16(localAt + 12, item.date, true)
    view.setUint32(localAt + 14, item.crc, true)
    view.setUint32(localAt + 18, item.payload.byteLength, true)
    view.setUint32(localAt + 22, item.size, true)
    view.setUint16(localAt + 26, item.nameBytes.byteLength, true)
    view.setUint16(localAt + 28, 0, true)
    out.set(item.nameBytes, localAt + LOCAL_HEADER_BYTES)
    out.set(item.payload, localAt + LOCAL_HEADER_BYTES + item.nameBytes.byteLength)

    view.setUint32(centralAt, CENTRAL_HEADER_SIGNATURE, true)
    view.setUint16(centralAt + 4, item.madeBy, true)
    view.setUint16(centralAt + 6, item.versionNeeded, true)
    view.setUint16(centralAt + 8, item.flags, true)
    view.setUint16(centralAt + 10, item.method, true)
    view.setUint16(centralAt + 12, item.time, true)
    view.setUint16(centralAt + 14, item.date, true)
    view.setUint32(centralAt + 16, item.crc, true)
    view.setUint32(centralAt + 20, item.payload.byteLength, true)
    view.setUint32(centralAt + 24, item.size, true)
    view.setUint16(centralAt + 28, item.nameBytes.byteLength, true)
    view.setUint16(centralAt + 30, 0, true)
    view.setUint16(centralAt + 32, 0, true)
    view.setUint16(centralAt + 34, 0, true)
    view.setUint16(centralAt + 36, item.internalAttributes, true)
    view.setUint32(centralAt + 38, item.externalAttributes, true)
    view.setUint32(centralAt + 42, localAt, true)
    out.set(item.nameBytes, centralAt + CENTRAL_HEADER_BYTES)

    localAt += LOCAL_HEADER_BYTES + item.nameBytes.byteLength + item.payload.byteLength
    centralAt += CENTRAL_HEADER_BYTES + item.nameBytes.byteLength
  }

  view.setUint32(centralAt, END_RECORD_SIGNATURE, true)
  view.setUint16(centralAt + 8, prepared.length, true)
  view.setUint16(centralAt + 10, prepared.length, true)
  view.setUint32(centralAt + 12, centralBytes, true)
  view.setUint32(centralAt + 16, centralOffset, true)
  return out
}

/** Everything a header needs for one entry: name bytes, flags, method, CRC, sizes and payload. */
function prepare(entry, encoder) {
  const name = entry.name
  if (typeof name !== 'string' || name.length === 0) throw new TypeError('zip entry without a name')
  const data = entry.data
  if (!(data instanceof Uint8Array)) throw new TypeError(`${name}: zip entry without data`)
  const nameBytes = encoder.encode(name)
  // Bit 3 says the sizes live in a data descriptor after the payload. They are written into the
  // header here, so it goes; every other general-purpose bit survives. Bit 11 is kept when the
  // file had it and added when the name only makes sense as UTF-8.
  let flags = (entry.flags ?? 0) & ~DATA_DESCRIPTOR_FLAG
  if (needsUtf8Flag(nameBytes)) flags |= UTF8_FLAG
  const method = entry.method === STORED ? STORED : DEFLATED
  const time = entry.time ?? 0
  const date = entry.date ?? 0x21
  const madeBy = entry.versionMadeBy ?? VERSION_20
  // The version a file says it needs is metadata a reader may act on; carrying the original value
  // keeps a rewrite byte-identical instead of stamping 2.0 over an archive that asked for less.
  const versionNeeded = entry.versionNeeded ?? VERSION_20
  const internalAttributes = entry.internalAttributes ?? 0
  const externalAttributes = entry.externalAttributes ?? 0
  const shared = {
    nameBytes,
    flags,
    method,
    time,
    date,
    madeBy,
    versionNeeded,
    internalAttributes,
    externalAttributes,
  }

  if (entry.changed !== true && entry.raw != null) {
    // Untouched: the compressed run, the CRC and the sizes the file had. The codec is not called,
    // so these payload bytes go out exactly as they came in. `data` must still be the content the
    // recorded CRC belongs to: copying a stale checksum into a new file would hide the damage.
    if (crc32(data) !== (entry.crc >>> 0)) throw new Error(`${name}: CRC mismatch`)
    const size = entry.uncompressedSize ?? data.byteLength
    return { ...shared, crc: entry.crc >>> 0, size, payload: entry.raw }
  }

  const payload = method === STORED ? data : deflateSync(data)
  return { ...shared, crc: crc32(data), size: data.byteLength, payload }
}

/** True when any name byte is outside ASCII, where readers would need the UTF-8 flag. */
function needsUtf8Flag(nameBytes) {
  for (let index = 0; index < nameBytes.length; index++) if (nameBytes[index] > 0x7f) return true
  return false
}

/**
 * The name of an entry from the central directory. Bit 11 means UTF-8; without it the field is
 * code page 437, unless a valid Info-ZIP Unicode Path extra field carries the real UTF-8 name.
 */
function decodeEntryName(nameBytes, flags, extraBytes) {
  if ((flags & UTF8_FLAG) !== 0) return UTF8_DECODER.decode(nameBytes)
  const unicode = unicodePathName(extraBytes, nameBytes)
  if (unicode !== null) return unicode
  return decodeCp437(nameBytes)
}

/**
 * The name from a 0x7075 extra field, or null when there is none (or its CRC does not match the
 * header name, which means the field and the record it annotates disagree).
 */
function unicodePathName(extra, originalName) {
  for (let at = 0; at + 4 <= extra.length; ) {
    const id = extra[at] | (extra[at + 1] << 8)
    const size = extra[at + 2] | (extra[at + 3] << 8)
    const body = at + 4
    if (body + size > extra.length) return null
    // Version 1, then the CRC of the original name field, then the UTF-8 name.
    if (id === EXTRA_UNICODE_PATH && size >= 6 && extra[body] === 1) {
      const expected =
        (extra[body + 1] | (extra[body + 2] << 8) | (extra[body + 3] << 16) | (extra[body + 4] << 24)) >>> 0
      if (expected === crc32(originalName)) return UTF8_DECODER.decode(extra.subarray(body + 5, body + size))
    }
    at = body + size
  }
  return null
}

/** Code page 437 text for the bytes of a name that carries no UTF-8 claim. */
function decodeCp437(bytes) {
  let text = ''
  for (let index = 0; index < bytes.length; index++) {
    const byte = bytes[index]
    text += String.fromCharCode(byte < 0x80 ? byte : CP437_HIGH[byte - 0x80])
  }
  return text
}

/**
 * Inflates `raw` and gives up as soon as the output passes `limit` bytes: the size a header
 * declares is not trusted, so a bomb is stopped by what actually comes out.
 */
function inflateWithin(raw, limit) {
  const chunks = []
  let produced = 0
  let overflow = false
  let broken = false
  const inflate = new Inflate((chunk) => {
    if (overflow) return
    produced += chunk.byteLength
    if (produced > limit) {
      overflow = true
      chunks.length = 0
      return
    }
    chunks.push(chunk)
  })
  const cap = Math.max(0, limit)
  try {
    if (raw.byteLength === 0) {
      inflate.push(EMPTY, true)
    } else {
      for (let at = 0; at < raw.byteLength; at += INFLATE_CHUNK_BYTES) {
        const stop = Math.min(at + INFLATE_CHUNK_BYTES, raw.byteLength)
        inflate.push(raw.subarray(at, stop), stop === raw.byteLength)
        if (overflow) break
      }
    }
  } catch (error) {
    // fflate throws "unexpected EOF" or "invalid block type" on damaged streams.
    broken = true
  }
  if (overflow) return { status: 'too-large', reason: `inflates past the ${cap} byte budget` }
  if (broken) return { status: 'corrupt', reason: 'the deflate stream is damaged' }

  const data = new Uint8Array(produced)
  let at = 0
  for (const chunk of chunks) {
    data.set(chunk, at)
    at += chunk.byteLength
  }
  return { status: 'ok', data }
}

/** Last end-of-central-directory record in the file, or -1. Scanned from the end, within 64 KiB. */
function findEndRecord(view) {
  const last = view.byteLength - END_RECORD_BYTES
  const first = Math.max(0, last - MAX_COMMENT_BYTES)
  let loose = -1
  for (let offset = last; offset >= first; offset--) {
    if (view.getUint32(offset, true) !== END_RECORD_SIGNATURE) continue
    // A record whose comment reaches exactly the end of the file is the real one; a signature
    // inside trailing junk is only a fallback.
    if (offset + END_RECORD_BYTES + view.getUint16(offset + 20, true) === view.byteLength) return offset
    if (loose < 0) loose = offset
  }
  return loose
}

function refuse(status, reason) {
  return { status, reason }
}
