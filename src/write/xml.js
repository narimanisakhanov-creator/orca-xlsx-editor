// Text-level XML surgery for OOXML parts.
//
// Nothing here ever re-serialises a document: every helper returns offsets into the
// original string so the caller can splice out only the bytes it means to change and
// leave the rest of the part byte-identical. A full DOM round-trip would normalise
// attribute order, quotes and empty-element form, which is exactly what "surgical
// save" must not do.

const WS = new Uint8Array(128);
for (const code of [9, 10, 13, 32]) WS[code] = 1;

function isWs(code) {
  return code < 128 && WS[code] === 1;
}

// XML 1.0 name characters, relaxed: anything non-ASCII is accepted because OOXML
// produced by Excel keeps ASCII names but locale-specific tools may not.
function isNameChar(code) {
  if (code < 128) {
    return (
      (code >= 65 && code <= 90) || // A-Z
      (code >= 97 && code <= 122) || // a-z
      (code >= 48 && code <= 57) || // 0-9
      code === 95 || // _
      code === 58 || // :
      code === 45 || // -
      code === 46 // .
    );
  }
  return true;
}

/** Local part of a possibly prefixed name: `x:row` -> `row`. */
export function localName(name) {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
}

/**
 * Scan one tag starting at `<` (index `lt`).
 * @returns {null | { kind: 'open'|'close'|'empty'|'other', name: string,
 *   start: number, tagEnd: number, attrsStart?: number, attrsEnd?: number }}
 */
function scanTag(xml, lt) {
  const len = xml.length;
  if (lt < 0 || lt >= len || xml.charCodeAt(lt) !== 60 /* < */) return null;
  let i = lt + 1;
  if (i >= len) return null;
  const first = xml.charCodeAt(i);

  if (first === 33 /* ! */) {
    if (xml.startsWith('<!--', lt)) {
      const end = xml.indexOf('-->', lt + 4);
      return end < 0 ? null : { kind: 'other', name: '', start: lt, tagEnd: end + 3 };
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt + 9);
      return end < 0 ? null : { kind: 'other', name: '', start: lt, tagEnd: end + 3 };
    }
    const end = xml.indexOf('>', lt);
    return end < 0 ? null : { kind: 'other', name: '', start: lt, tagEnd: end + 1 };
  }

  if (first === 63 /* ? */) {
    const end = xml.indexOf('?>', lt);
    return end < 0 ? null : { kind: 'other', name: '', start: lt, tagEnd: end + 2 };
  }

  if (first === 47 /* / */) {
    let j = i + 1;
    while (j < len && isNameChar(xml.charCodeAt(j))) j++;
    const name = xml.slice(i + 1, j);
    const end = xml.indexOf('>', j);
    return end < 0 ? null : { kind: 'close', name, start: lt, tagEnd: end + 1 };
  }

  let j = i;
  while (j < len && isNameChar(xml.charCodeAt(j))) j++;
  if (j === i) return null;
  const name = xml.slice(i, j);
  const attrsStart = j;

  // Attribute values may legally contain `>`, so quote-aware scanning is required.
  let k = j;
  while (k < len) {
    const ch = xml[k];
    if (ch === '"' || ch === "'") {
      const close = xml.indexOf(ch, k + 1);
      if (close < 0) return null;
      k = close + 1;
      continue;
    }
    if (ch === '>') break;
    k++;
  }
  if (k >= len) return null;

  let t = k - 1;
  while (t > attrsStart && isWs(xml.charCodeAt(t))) t--;
  const selfClosing = xml.charCodeAt(t) === 47 /* / */;
  if (selfClosing) {
    t--;
    while (t > attrsStart && isWs(xml.charCodeAt(t))) t--;
  }
  const attrsEnd = t + 1;
  return {
    kind: selfClosing ? 'empty' : 'open',
    name,
    start: lt,
    tagEnd: k + 1,
    attrsStart,
    attrsEnd,
  };
}

function matchElement(xml, tok, wanted) {
  const attrs = xml.slice(tok.attrsStart, tok.attrsEnd);
  if (tok.kind === 'empty') {
    return {
      name: tok.name,
      start: tok.start,
      innerStart: tok.tagEnd,
      innerEnd: tok.tagEnd,
      end: tok.tagEnd,
      attrs,
      selfClosing: true,
      tagEnd: tok.tagEnd,
    };
  }
  let depth = 1;
  let i = tok.tagEnd;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    const t = scanTag(xml, lt);
    if (!t) break;
    if (t.kind === 'open' && localName(t.name) === wanted) depth++;
    else if (t.kind === 'empty' && localName(t.name) === wanted) {
      // self-closing nested element never closes a parent
    } else if (t.kind === 'close' && localName(t.name) === wanted) {
      depth--;
      if (depth === 0) {
        return {
          name: tok.name,
          start: tok.start,
          innerStart: tok.tagEnd,
          innerEnd: lt,
          end: t.tagEnd,
          attrs,
          selfClosing: false,
          tagEnd: tok.tagEnd,
        };
      }
    }
    i = t.tagEnd;
  }
  return null;
}

/**
 * All elements with the given local name, in document order.
 * Offsets refer to `xml`; `attrs` is the raw attribute text of the start tag.
 */
export function findElements(xml, name) {
  const out = [];
  if (typeof xml !== 'string' || typeof name !== 'string' || name === '') return out;
  let i = xml.indexOf('<');
  while (i >= 0) {
    const tok = scanTag(xml, i);
    if (!tok) break;
    if ((tok.kind === 'open' || tok.kind === 'empty') && localName(tok.name) === name) {
      const el = matchElement(xml, tok, name);
      if (el) {
        out.push(el);
        i = xml.indexOf('<', el.end);
        continue;
      }
    }
    i = xml.indexOf('<', tok.tagEnd);
  }
  return out;
}

/** First element with that local name, or null. */
export function findElement(xml, name) {
  return findElements(xml, name)[0] ?? null;
}

function firstElementToken(xml) {
  if (typeof xml !== 'string') return null;
  let i = xml.indexOf('<');
  while (i >= 0) {
    const tok = scanTag(xml, i);
    if (!tok) return null;
    if (tok.kind === 'open' || tok.kind === 'empty') return tok;
    i = xml.indexOf('<', tok.tagEnd);
  }
  return null;
}

/** Root element local name, e.g. `worksheet`; null when there is no element. */
export function rootName(xml) {
  const tok = firstElementToken(xml);
  return tok ? localName(tok.name) : null;
}

/** Namespace prefix of the root element, `x:` for `<x:worksheet>`; empty when unprefixed. */
export function rootPrefix(xml) {
  const tok = firstElementToken(xml);
  if (!tok) return '';
  const i = tok.name.indexOf(':');
  return i < 0 ? '' : tok.name.slice(0, i + 1);
}

/** Local names of the root element's direct children, in document order. */
export function topLevelNames(xml) {
  return topLevelElements(xml).map((el) => el.local);
}

/**
 * Direct children of the root element, with offsets. Nested elements that happen to
 * share a name (an `extLst` inside `sheetPr`, say) are excluded, which is the
 * difference between "insert before the doc-level extLst" and corrupting the XML.
 */
export function topLevelElements(xml) {
  const root = firstElementToken(xml);
  if (!root || root.kind === 'empty') return [];
  const out = [];
  let depth = 1;
  let i = root.tagEnd;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    const tok = scanTag(xml, lt);
    if (!tok) break;
    if (tok.kind === 'open' || tok.kind === 'empty') {
      if (depth === 1) {
        const el = matchElement(xml, tok, localName(tok.name));
        if (el && !el.selfClosing) {
          out.push({ ...el, local: localName(tok.name) });
          i = el.end;
          continue;
        }
        if (el) out.push({ ...el, local: localName(tok.name) });
      } else if (tok.kind === 'open') {
        depth++;
      }
    } else if (tok.kind === 'close') {
      depth--;
      if (depth === 0) break;
    }
    i = tok.tagEnd;
  }
  return out;
}

/** Attribute records of a start tag's raw attribute text. */
function attrRecords(attrsText) {
  const records = [];
  const n = attrsText.length;
  let i = 0;
  while (i < n) {
    while (i < n && isWs(attrsText.charCodeAt(i))) i++;
    if (i >= n) break;
    const nameStart = i;
    while (i < n && isNameChar(attrsText.charCodeAt(i))) i++;
    if (i === nameStart) {
      i++;
      continue;
    }
    const name = attrsText.slice(nameStart, i);
    while (i < n && isWs(attrsText.charCodeAt(i))) i++;
    if (attrsText[i] !== '=') {
      records.push({ name, value: '', start: nameStart, end: i, valueStart: i, valueEnd: i });
      continue;
    }
    i++;
    while (i < n && isWs(attrsText.charCodeAt(i))) i++;
    const quote = attrsText[i];
    let valueStart;
    let valueEnd;
    if (quote === '"' || quote === "'") {
      valueStart = i + 1;
      const close = attrsText.indexOf(quote, valueStart);
      valueEnd = close < 0 ? n : close;
      i = close < 0 ? n : close + 1;
    } else {
      valueStart = i;
      while (i < n && !isWs(attrsText.charCodeAt(i))) i++;
      valueEnd = i;
    }
    records.push({
      name,
      value: attrsText.slice(valueStart, valueEnd),
      start: nameStart,
      end: i,
      valueStart,
      valueEnd,
    });
  }
  return records;
}

/**
 * Attribute value, unescaped. Prefixed lookups (`r:id`) fall back to a local-name
 * match so a package that declares the relationship namespace under another prefix
 * still resolves.
 */
export function getAttr(attrsText, name) {
  if (typeof attrsText !== 'string' || typeof name !== 'string') return null;
  const records = attrRecords(attrsText);
  const exact = records.find((r) => r.name === name);
  if (exact) return unescapeXml(exact.value);
  if (name.includes(':')) {
    const wanted = localName(name);
    const loose = records.find((r) => r.name.includes(':') && localName(r.name) === wanted);
    if (loose) return unescapeXml(loose.value);
  }
  return null;
}

/** Replace or append an attribute in raw attribute text. */
export function setAttr(attrsText, name, value) {
  const text = typeof attrsText === 'string' ? attrsText : '';
  const escaped = escapeXml(String(value));
  const records = attrRecords(text);
  const existing = records.find((r) => r.name === name);
  if (existing) {
    return text.slice(0, existing.valueStart) + escaped + text.slice(existing.valueEnd);
  }
  return `${text} ${name}="${escaped}"`;
}

/** Remove an attribute (with the whitespace that precedes it) from raw attribute text. */
export function removeAttr(attrsText, name) {
  if (typeof attrsText !== 'string') return '';
  const record = attrRecords(attrsText).find((r) => r.name === name);
  if (!record) return attrsText;
  let from = record.start;
  while (from > 0 && isWs(attrsText.charCodeAt(from - 1))) from--;
  return attrsText.slice(0, from) + attrsText.slice(record.end);
}

/**
 * Replace each matching element. `replacer(element, elementText)` returns the new
 * text for that element, or null/undefined to leave it alone. Splices run back to
 * front so earlier offsets stay valid even when lengths change.
 */
export function replaceElements(xml, name, replacer) {
  const elements = findElements(xml, name);
  let out = xml;
  for (let i = elements.length - 1; i >= 0; i--) {
    const el = elements[i];
    const current = out.slice(el.start, el.end);
    const next = replacer(el, current);
    if (next == null || next === current) continue;
    out = out.slice(0, el.start) + next + out.slice(el.end);
  }
  return out;
}

const ESCAPE_TEST = /[&<>"'\u0000-\u0008\u000B\u000C\u000E-\u001F]/;
const NAMED = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** Escape text for element content and attribute values (invisible control chars are dropped). */
export function escapeXml(text) {
  const str = String(text);
  if (!ESCAPE_TEST.test(str)) return str;
  let out = '';
  for (const ch of str) {
    switch (ch) {
      case '&':
        out += '&amp;';
        break;
      case '<':
        out += '&lt;';
        break;
      case '>':
        out += '&gt;';
        break;
      case '"':
        out += '&quot;';
        break;
      case "'":
        out += '&apos;';
        break;
      default: {
        const code = ch.codePointAt(0);
        // XML 1.0 forbids most C0 controls; keeping them would make the part unparsable.
        if (code < 0x20 && code !== 9 && code !== 10 && code !== 13) break;
        out += ch;
      }
    }
  }
  return out;
}

/** Decode entities produced by escapeXml plus numeric character references. */
export function unescapeXml(text) {
  const str = String(text);
  if (!str.includes('&')) return str;
  return str.replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|([a-zA-Z]+));/g, (whole, dec, hex, named) => {
    if (dec != null || hex != null) {
      const code = dec != null ? Number(dec) : parseInt(hex, 16);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return whole;
      if (code >= 0xd800 && code <= 0xdfff) return whole;
      return String.fromCodePoint(code);
    }
    const mapped = NAMED[named];
    return mapped === undefined ? whole : mapped;
  });
}
