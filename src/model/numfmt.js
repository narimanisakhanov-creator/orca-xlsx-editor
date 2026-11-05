// Number formats to display text. A workbook read by a user has to look like it does in Excel, so
// this covers the codes that actually appear in working files: General, fixed and grouped numbers,
// percentages, scientific, currency literals, dates and times, and `;`-separated sections.
//
// Anything that is not understood falls back to General: the value is never dropped or replaced by
// an error, a wrong-looking number is better than a blank cell.
//
// Deliberately no locale: month and weekday names are English, decimal separator is '.', group
// separator is ','. Excel resolves those from the system locale, which a workbook does not carry.

const DAY_MS = 86400000
// Days between an Excel epoch and the Unix epoch. 1899-12-30 is the 1900 system (the same constant
// ExcelJS uses); 1904-01-01 is the 1904 system, 1462 days later.
const EXCEL_EPOCH_OFFSET = 25569
const EXCEL_EPOCH_1904 = 24107

// Characters a format code may carry inside the integer core as a thousands separator. Excel
// prints them at the position the pattern puts them in — `#\ ##0` on 1234567 is `1234 567`, not
// `1 234 567` — which is what `renderInteger` reproduces. The capture keeps the separators in the
// split that cuts the pattern into digit groups.
const GROUP_SEPARATOR = /([\u0020\u00a0'])/

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const WEEKDAYS_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

const NUMERIC_PLACEHOLDERS = '0#?,.'

/**
 * True when a format code renders a date or a time. The first section decides, since that is the
 * one a date is rendered with; quotes and brackets are stripped from it first, so
 * `#,##0.00 "RUB"` and `[h]:mm:ss` are classified by what is left.
 */
export function isDateFormat(numFmt) {
  if (typeof numFmt !== 'string') return false
  const trimmed = numFmt.trim()
  if (trimmed === '' || /^general$/i.test(trimmed)) return false
  // Only the section a date would be rendered with decides: `[$-419]d\ mmm;@` is a date format, and
  // a text section such as `;@` must not be scanned for date letters.
  const first = splitSections(trimmed)[0] ?? ''
  const bare = first.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')
  return /[ymdhs]/i.test(bare)
}

/**
 * Formatted text for one cell value.
 *
 * `isDate` is the caller's own verdict (`value` is a Date or the cell's format is a date format);
 * it is passed in so a formula result that ExcelJS typed as a number still renders as a date.
 * `date1904` selects the workbook's epoch: a serial in a 1904 book counts from 1904-01-01, so the
 * same number is 1462 days later than in a 1900 book.
 */
export function formatValue(value, numFmt, isDate = false, date1904 = false) {
  if (value === null || value === undefined) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'

  const fmt = typeof numFmt === 'string' && numFmt.trim() !== '' ? numFmt : 'General'

  if (value instanceof Date) {
    const serial = value.getTime() / DAY_MS + (date1904 ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_OFFSET)
    return isDateFormat(fmt) ? formatDate(serial, fmt, date1904) : general(serial)
  }
  if (typeof value === 'string') return formatText(value, fmt)
  if (typeof value !== 'number' || !Number.isFinite(value)) return String(value)
  if (isDate || isDateFormat(fmt)) return formatDate(value, fmt, date1904)
  return formatNumber(value, fmt)
}

/** Value as the grid should show it: unformatted numbers, so the caller can reuse this. */
function general(value) {
  if (!Number.isFinite(value)) return String(value)
  if (value === 0) return '0'
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value)

  const magnitude = Math.abs(value)
  if (magnitude >= 1e11 || magnitude < 1e-10) {
    const [mantissa, exponent] = value.toExponential(5).split('e')
    const digits = mantissa.replace(/0+$/, '').replace(/\.$/, '')
    const sign = exponent.startsWith('-') ? '-' : '+'
    return `${digits}E${sign}${exponent.replace('-', '').padStart(2, '0')}`
  }
  // 11 significant digits is Excel's General width; it also hides the 0.08119999999999999 kind of
  // noise that binary floating point leaves behind.
  return String(parseFloat(value.toPrecision(11)))
}

function formatNumber(value, fmt) {
  if (/^\s*general\s*$/i.test(fmt)) return general(value)

  const sections = splitSections(fmt)
  const conditions = sections.map(sectionCondition)
  let section
  let explicitNegative = false

  if (conditions.some(Boolean)) {
    // Conditional sections replace the sign rule: the first section whose test the value passes
    // wins, and a section with no test is the fallback. Excel shows the plain number when neither
    // exists (`[>3]0` on 1 is `1`).
    let index = conditions.findIndex((condition) => condition !== null && conditionHolds(value, condition))
    if (index === -1) index = conditions.indexOf(null)
    if (index === -1) return general(value)
    section = sections[index]
    explicitNegative = value < 0
  } else if (value < 0 && sections.length > 1 && sections[1].trim() !== '') {
    section = sections[1]
    explicitNegative = true
  } else if (value === 0 && sections.length > 2 && sections[2].trim() !== '') {
    section = sections[2]
  } else {
    section = sections[0]
  }

  const rendered = section === undefined ? null : renderSection(Math.abs(value), section)
  if (rendered === null) return general(value)
  // A literal-only section says what to print, sign included (`;\–`), so nothing is added to it.
  if (value >= 0 || explicitNegative || rendered.plain || /^\s*-/.test(rendered.text)) return rendered.text
  return `-${rendered.text}`
}

/**
 * The test a section starts with (`[>=1000000]`, `[>1000]`), or `null`. Colour and locale
 * brackets (`[Red]`, `[$-419]`) may precede it and are skipped.
 */
function sectionCondition(section) {
  let at = 0
  while (section[at] === '[') {
    const end = section.indexOf(']', at)
    if (end === -1) return null
    const inner = section.slice(at + 1, end).trim()
    const match = /^(>=|<=|<>|=|>|<)\s*(-?\d+(?:\.\d+)?)$/.exec(inner)
    if (match) return { operator: match[1], operand: Number(match[2]) }
    at = end + 1
  }
  return null
}

function conditionHolds(value, condition) {
  switch (condition.operator) {
    case '>': return value > condition.operand
    case '<': return value < condition.operand
    case '>=': return value >= condition.operand
    case '<=': return value <= condition.operand
    case '=': return value === condition.operand
    default: return value !== condition.operand
  }
}

/** Split on `;` only outside quotes and brackets: `[Red];` and `"a;b"` must not split. */
function splitSections(fmt) {
  const sections = []
  let current = ''
  let quoted = false
  let brackets = 0
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]
    if (ch === '"') {
      quoted = !quoted
      current += ch
      continue
    }
    if (!quoted) {
      if (ch === '[') brackets++
      else if (ch === ']') brackets = Math.max(0, brackets - 1)
      else if (ch === '\\') {
        current += ch + (fmt[i + 1] ?? '')
        i++
        continue
      }
      if (ch === ';' && brackets === 0) {
        sections.push(current)
        current = ''
        continue
      }
    }
    current += ch
  }
  sections.push(current)
  return sections
}

/**
 * Turn one section into typed tokens. Literal characters (quoted text, currency symbols, escapes)
 * end up as `lit`, placeholders as `p`; brackets that are not currency (colours, conditions)
 * vanish, which is what Excel shows.
 */
function tokenize(section) {
  const tokens = []
  let i = 0
  while (i < section.length) {
    const ch = section[i]

    if (ch === '"') {
      let j = i + 1
      while (j < section.length && section[j] !== '"') j++
      for (const char of section.slice(i + 1, j)) tokens.push({ ...literal(char), escaped: true })
      i = j + 1
      continue
    }

    if (ch === '[') {
      const end = section.indexOf(']', i)
      const inner = end === -1 ? '' : section.slice(i + 1, end)
      i = end === -1 ? section.length : end + 1
      if (inner.startsWith('$')) {
        const body = inner.slice(1)
        const dash = body.indexOf('-')
        const symbol = (dash === -1 ? body : body.slice(0, dash)).replace(/"/g, '')
        for (const char of symbol) tokens.push({ lit: char })
      }
      continue
    }

    if (ch === '\\') {
      if (i + 1 < section.length) tokens.push({ ...literal(section[i + 1]), escaped: true })
      i += 2
      continue
    }
    if (ch === '_') {
      // Width placeholder: Excel reserves the width of the next character and prints nothing. It is
      // padding, not a literal, so it never acts as a thousands separator.
      tokens.push({ lit: ' ', escaped: true })
      i += 2
      continue
    }
    if (ch === '*') {
      i += 2
      continue
    }
    if (ch === '%') {
      tokens.push({ lit: '%', percent: true })
      i++
      continue
    }
    if (NUMERIC_PLACEHOLDERS.includes(ch)) {
      tokens.push({ p: ch })
      i++
      continue
    }
    if ((ch === 'E' || ch === 'e') && (section[i + 1] === '+' || section[i + 1] === '-') && '0#'.includes(section[i + 2] ?? '')) {
      tokens.push({ p: 'E', expSign: section[i + 1] })
      i += 2
      continue
    }
    if (ch === '@') {
      tokens.push({ text: true })
      i++
      continue
    }
    tokens.push(literal(ch))
    i++
  }
  return tokens
}

/**
 * A literal token. `separator` marks the characters that group digits inside the integer core
 * (`#\ ##0`, `###\ ###\ ###`); every other literal inside the core is decoration and is dropped,
 * which is how `renderSection` has always treated them.
 */
function literal(char) {
  return GROUP_SEPARATOR.test(char) ? { lit: char, separator: true } : { lit: char }
}

/**
 * One section to its text, or `null` when the section is not a format this module understands (the
 * caller then falls back to General). `plain` marks a section with no numeric placeholder: its text
 * is printed as written, sign included, so no sign is added to it.
 */
function renderSection(value, section) {
  const tokens = tokenize(section)
  let first = -1
  let last = -1
  for (let i = 0; i < tokens.length; i++) {
    if (tokens[i].p) {
      if (first === -1) first = i
      last = i
    }
  }
  // A section with no placeholder is a literal (`"zero"`, `\–`) and Excel prints it as written for
  // that sign. Plain characters with no escaping (`zzz`) are an unknown code, not a literal.
  if (first === -1) {
    const literalOnly = tokens.every((token) => token.escaped === true)
    return literalOnly ? { text: literalText(tokens), plain: true } : null
  }

  const prefix = literalText(tokens.slice(0, first))
  const suffix = literalText(tokens.slice(last + 1))
  // The core keeps only placeholders and group separators: any other literal between them is
  // decoration that the pattern does not render (`#\ x##0` prints digits only). A separator after
  // the decimal point is decoration too — there are no digit groups to separate there.
  let core = ''
  let expSign = '+'
  let inFraction = false
  for (let i = first; i <= last; i++) {
    const token = tokens[i]
    if (token.p) {
      core += token.p
      if (token.p === '.') inFraction = true
      if (token.expSign) expSign = token.expSign
    } else if (token.separator && !inFraction) {
      core += token.lit
    }
  }
  if (core === '') return null

  const percent = tokens.some((token) => token.percent)
  const number = renderCore(core, percent ? value * 100 : value, expSign)
  if (number === null) return null
  return { text: `${prefix}${number}${suffix}`, plain: false }
}

function literalText(tokens) {
  let text = ''
  for (const token of tokens) {
    if (token.text) text += '@'
    else if (token.lit !== undefined) text += token.lit
  }
  return text
}

function renderCore(core, value, expSign) {
  // Commas after the last placeholder scale the number instead of separating digits: each one
  // divides by 1000 and none becomes a decimal (`#,##0.0,," млн"` shows 1234567 as `1.2`).
  let mantissa = core
  let scale = 0
  const scaling = /,+$/.exec(mantissa)
  if (scaling) {
    scale = scaling[0].length
    mantissa = mantissa.slice(0, -scale)
  }
  if (scale > 0) value /= 1000 ** scale

  const sci = /([E])([+-]?)(0+)$/.exec(mantissa)
  let exponentDigits = 0
  let exponentSign = expSign
  if (sci) {
    mantissa = mantissa.slice(0, sci.index)
    exponentDigits = sci[3].length
    exponentSign = sci[2] || expSign
  }

  const [intPattern = '', decPattern = ''] = mantissa.split('.')
  const decimals = decPattern.length
  const optionalDecimals = (decPattern.match(/[#?]/g) ?? []).length

  if (sci) {
    let exponent = 0
    let scaled = value
    if (value > 0) {
      exponent = Math.floor(Math.log10(value))
      scaled = value / 10 ** exponent
      if (scaled >= 10) {
        scaled /= 10
        exponent++
      } else if (scaled < 1) {
        scaled *= 10
        exponent--
      }
    }
    const digits = fixed(scaled, Math.max(decimals, 1))
    const sign = exponent < 0 ? '-' : exponentSign === '+' ? '+' : ''
    return `${digits}E${sign}${String(Math.abs(exponent)).padStart(exponentDigits, '0')}`
  }

  const text = fixed(value, decimals)
  const [integerDigits, rawFraction = ''] = text.split('.')
  let fractionPart = rawFraction
  if (optionalDecimals > 0) {
    // `#` and `?` may drop trailing zeros; they must leave the zeros that `0` placeholders demand.
    const keep = Math.max(0, decimals - optionalDecimals)
    while (fractionPart.length > keep && fractionPart.endsWith('0')) fractionPart = fractionPart.slice(0, -1)
  }
  // `?` reserves the width of a digit it did not print, which is what aligns columns in the
  // accounting formats.
  if (decPattern.includes('?') && fractionPart.length < decimals) fractionPart = fractionPart.padEnd(decimals, ' ')

  const integerPart = renderInteger(intPattern, integerDigits)
  return fractionPart === '' ? integerPart : `${integerPart}.${fractionPart}`
}

/**
 * The integer digits for one pattern.
 *
 * A comma groups every three digits and keeps repeating as far as the number goes (`#,##0`).
 * A literal separator — space, no-break space, apostrophe — instead holds the position the pattern
 * gives it and the digits that do not fit are prepended, which is how Excel reads `#\ ##0`: 1234 is
 * `1 234` and 1234567 is `1234 567`, not `1 234 567`.
 */
function renderInteger(pattern, digits) {
  // Neither `#` nor `?` requires a zero, so the single 0 a rounded value leaves behind disappears.
  if (digits === '0' && !pattern.includes('0')) digits = ''

  if (pattern.includes(',')) {
    return renderGroup(pattern.replace(/,/g, ''), digits).replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  }

  const parts = pattern.split(GROUP_SEPARATOR)
  if (parts.length === 1) return renderGroup(pattern, digits)

  const groups = []
  const separators = []
  for (let i = 0; i < parts.length; i++) {
    if (i % 2 === 0) groups.push(parts[i])
    else separators.push(parts[i])
  }

  const rendered = new Array(groups.length)
  let rest = digits
  for (let i = groups.length - 1; i >= 0; i--) {
    const take = i === 0 ? rest.length : Math.min(groups[i].length, rest.length)
    rendered[i] = renderGroup(groups[i], rest.slice(rest.length - take))
    rest = rest.slice(0, rest.length - take)
  }

  let out = rendered[0]
  for (let i = 1; i < rendered.length; i++) out += separators[i - 1] + rendered[i]
  return out
}

/**
 * One group of digit placeholders, filled from the right. `0` always prints a digit, `#` prints one
 * only when there is one, `?` prints a space where it has no digit; digits past the pattern are
 * prepended.
 */
function renderGroup(pattern, digits) {
  let out = ''
  let at = digits.length - 1
  for (let i = pattern.length - 1; i >= 0; i--) {
    if (at >= 0) {
      out = digits[at] + out
      at--
    } else if (pattern[i] === '0') {
      out = `0${out}`
    } else if (pattern[i] === '?') {
      out = ` ${out}`
    }
  }
  return at >= 0 ? digits.slice(0, at + 1) + out : out
}

/**
 * Round to `decimals` places on a non-negative value.
 *
 * `toFixed` rounds on the binary value, so a decimal tie like 2.675 can land on 2.67 where Excel
 * prints 2.68. Ties are rare in the numbers a workbook displays and the error is one unit in the
 * last place; matching Excel exactly would mean decimal arithmetic for every cell, which is not
 * worth the cost on the render path.
 */
function fixed(value, decimals) {
  return decimals <= 0 ? String(Math.round(value)) : value.toFixed(decimals)
}

/* ------------------------------------------------------------------ dates ---------------------- */

function formatDate(serial, fmt, date1904 = false) {
  if (!isDateFormat(fmt)) return general(serial)

  // A date is rendered by the first section: `d mmm;@` uses `d mmm` for a date and `@` for text,
  // and the `;` is a section break, never a literal.
  const section = splitSections(fmt)[0] ?? fmt
  const epoch = date1904 ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_OFFSET
  const date = new Date(Math.round((serial - epoch) * DAY_MS))
  const tokens = dateTokens(section)
  const usesAmPm = tokens.some((token) => token.kind === 'ampm')
  const totalSeconds = Math.round(Math.abs(serial) * 86400)

  let out = ''
  let previousField = null
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]
    if (token.kind === 'literal') {
      out += token.text
      continue
    }
    if (token.kind === 'ampm') {
      out += date.getUTCHours() < 12 ? (token.short ? 'A' : 'AM') : (token.short ? 'P' : 'PM')
      previousField = 'ampm'
      continue
    }
    if (token.kind === 'elapsed') {
      if (token.unit === 'h') out += String(Math.floor(totalSeconds / 3600)).padStart(token.run, '0')
      else if (token.unit === 'm') out += String(Math.floor(totalSeconds / 60) % 60).padStart(token.run, '0')
      else out += String(totalSeconds % 60).padStart(token.run, '0')
      previousField = token.unit === 'h' ? 'hour' : token.unit === 'm' ? 'minute' : 'second'
      continue
    }

    const field = resolveField(token, tokens, i, previousField)
    out += renderField(field, token.run, date, usesAmPm)
    previousField = field
  }
  return out
}

function dateTokens(fmt) {
  const tokens = []
  const lower = fmt.toLowerCase()
  let i = 0
  while (i < fmt.length) {
    const ch = fmt[i]

    if (ch === '"') {
      let j = i + 1
      while (j < fmt.length && fmt[j] !== '"') j++
      tokens.push({ kind: 'literal', text: fmt.slice(i + 1, j) })
      i = j + 1
      continue
    }
    if (ch === '\\') {
      tokens.push({ kind: 'literal', text: fmt[i + 1] ?? '' })
      i += 2
      continue
    }
    if (ch === '[') {
      const end = fmt.indexOf(']', i)
      const inner = lower.slice(i + 1, end === -1 ? fmt.length : end)
      i = end === -1 ? fmt.length : end + 1
      if (/^[hms]+$/.test(inner)) tokens.push({ kind: 'elapsed', unit: inner[0], run: inner.length })
      else tokens.push({ kind: 'literal', text: '' })
      continue
    }
    if (lower.startsWith('am/pm', i)) {
      tokens.push({ kind: 'ampm' })
      i += 5
      continue
    }
    if (lower.startsWith('a/p', i)) {
      tokens.push({ kind: 'ampm', short: true })
      i += 3
      continue
    }
    if (ch === '_' || ch === '*') {
      i += 2
      continue
    }
    if ('ymdhs'.includes(lower[i])) {
      let j = i
      while (j < fmt.length && lower[j] === lower[i]) j++
      tokens.push({ kind: 'letter', letter: lower[i], run: j - i })
      i = j
      continue
    }
    tokens.push({ kind: 'literal', text: ch })
    i++
  }
  return tokens
}

function resolveField(token, tokens, index, previousField) {
  switch (token.letter) {
    case 'y':
      return 'year'
    case 'd':
      return 'day'
    case 'h':
      return 'hour'
    case 's':
      return 'second'
    case 'm': {
      if (token.run >= 3) return 'month'
      // `m` is a minute next to an hour or a second, a month everywhere else: `hh:mm`, `mm:ss`.
      if (previousField === 'hour') return 'minute'
      return nextLetter(tokens, index) === 's' ? 'minute' : 'month'
    }
    default:
      return 'literal'
  }
}

function nextLetter(tokens, index) {
  for (let i = index + 1; i < tokens.length; i++) {
    if (tokens[i].kind === 'letter') return tokens[i].letter
    if (tokens[i].kind !== 'literal') return null
  }
  return null
}

function renderField(field, run, date, usesAmPm) {
  switch (field) {
    case 'year':
      return run <= 2 ? String(date.getUTCFullYear() % 100).padStart(2, '0') : String(date.getUTCFullYear())
    case 'month': {
      const month = date.getUTCMonth()
      if (run === 1) return String(month + 1)
      if (run === 2) return String(month + 1).padStart(2, '0')
      if (run === 3) return MONTHS_SHORT[month]
      if (run === 4) return MONTHS[month]
      return MONTHS[month][0]
    }
    case 'day': {
      const day = date.getUTCDate()
      if (run === 1) return String(day)
      if (run === 2) return String(day).padStart(2, '0')
      if (run === 3) return WEEKDAYS_SHORT[date.getUTCDay()]
      return WEEKDAYS[date.getUTCDay()]
    }
    case 'hour': {
      const hours = date.getUTCHours()
      const shown = usesAmPm ? (hours % 12 === 0 ? 12 : hours % 12) : hours
      return String(shown).padStart(run >= 2 ? 2 : 1, '0')
    }
    case 'minute':
      return String(date.getUTCMinutes()).padStart(run >= 2 ? 2 : 1, '0')
    case 'second':
      return String(date.getUTCSeconds()).padStart(run >= 2 ? 2 : 1, '0')
    default:
      return ''
  }
}

/* ------------------------------------------------------------------ text ----------------------- */

/**
 * The section that carries `@` renders a text value: `"n/a";;;@` prints the text, `d mmm;@` prints
 * it unchanged. With no `@` anywhere the text is shown as it is, which is what Excel does.
 */
function formatText(text, fmt) {
  if (!fmt.includes('@')) return text
  const sections = splitSections(fmt)
  for (let i = sections.length - 1; i >= 0; i--) {
    const tokens = tokenize(sections[i])
    if (tokens.some((token) => token.text)) return tokens.map((token) => (token.text ? text : token.lit ?? '')).join('')
  }
  return text
}
