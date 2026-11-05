// ExcelJS style objects to CSS-ready values. Nothing here touches the DOM: the grid decides how to
// apply a colour, this module only resolves what the colour is.
//
// Colours arrive as `argb` (alpha is dropped: a cell fill is opaque), as a `theme` index plus an
// optional `tint`, or as a legacy `indexed` palette entry. Theme indices only mean something with
// the workbook's theme part, so the caller passes the palette built by `themePalette`.

// Excel indexes theme colours in this order, which is not the order the theme XML lists them in.
const THEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

// The first 16 entries of the legacy indexed palette; the rest depend on the workbook.
const INDEXED_COLORS = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'
]

/**
 * The twelve theme colours of a workbook, in Excel's index order, or `null` when the theme part is
 * missing (a workbook may legitimately have none). Each entry is `#RRGGBB`.
 */
export function themePalette(themes) {
  const xml = themes && typeof themes === 'object' ? Object.values(themes)[0] : null
  if (typeof xml !== 'string') return null

  const scheme = /<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(xml)
  if (!scheme) return null

  const byName = new Map()
  const element = /<a:(\w+)>([\s\S]*?)<\/a:\1>/g
  let match
  while ((match = element.exec(scheme[1])) !== null) {
    const [, name, body] = match
    const srgb = /<a:srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(body)
    if (srgb) {
      byName.set(name, `#${srgb[1].toUpperCase()}`)
      continue
    }
    const sys = /<a:sysClr\b[^>]*\b(?:lastClr|val)="([0-9A-Fa-f]{6}|[A-Za-z]+)"/.exec(body)
    if (sys) byName.set(name, sysColor(sys[1]))
  }

  const palette = THEME_ORDER.map((name) => byName.get(name) ?? null)
  return palette.some(Boolean) ? palette : null
}

function sysColor(value) {
  if (/^[0-9A-Fa-f]{6}$/.test(value)) return `#${value.toUpperCase()}`
  switch (value) {
    case 'window':
    case 'windowBackground':
      return '#FFFFFF'
    case 'windowText':
    case 'windowFrame':
    case 'captionText':
    case 'menuText':
      return '#000000'
    default:
      return null
  }
}

/** `#RRGGBB` for one ExcelJS colour, or `null` when it cannot be resolved. */
export function cssColor(color, palette, notes) {
  if (!color) return null

  if (typeof color.argb === 'string' && color.argb.length >= 6) {
    // ARGB keeps its alpha for charts; a worksheet colour is opaque, so AA is dropped.
    return `#${color.argb.slice(-6).toUpperCase()}`
  }

  if (typeof color.theme === 'number') {
    const base = palette?.[color.theme]
    if (!base) {
      notes?.push(`theme colour ${color.theme} is not in the workbook theme, left unset`)
      return null
    }
    return typeof color.tint === 'number' && color.tint !== 0 ? applyTint(base, color.tint) : base
  }

  if (typeof color.indexed === 'number') {
    const base = INDEXED_COLORS[color.indexed]
    if (base) return base
    notes?.push(`indexed colour ${color.indexed} is outside the built-in palette, left unset`)
    return null
  }

  return null
}

/**
 * Excel's tint: a positive tint lightens towards white, a negative one darkens towards black, on
 * the HSL lightness channel. `tint` is the -1..1 fraction stored in the style.
 */
export function applyTint(hex, tint) {
  const [h, s, l] = rgbToHsl(hex)
  const clamped = Math.max(-1, Math.min(1, tint))
  const shifted = clamped < 0 ? l * (1 + clamped) : l * (1 - clamped) + clamped
  return hslToHex(h, s, shifted)
}

function rgbToHsl(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255
  const g = parseInt(hex.slice(3, 5), 16) / 255
  const b = parseInt(hex.slice(5, 7), 16) / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]

  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h
  if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6
  else if (max === g) h = ((b - r) / d + 2) / 6
  else h = ((r - g) / d + 4) / 6
  return [h, s, l]
}

function hslToHex(h, s, l) {
  const channel = (p, q, t) => {
    let value = t
    if (value < 0) value += 1
    if (value > 1) value -= 1
    if (value < 1 / 6) return p + (q - p) * 6 * value
    if (value < 1 / 2) return q
    if (value < 2 / 3) return p + (q - p) * (2 / 3 - value) * 6
    return p
  }

  let r
  let g
  let b
  if (s === 0) {
    r = g = b = l
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    r = channel(p, q, h + 1 / 3)
    g = channel(p, q, h)
    b = channel(p, q, h - 1 / 3)
  }
  const byte = (value) => Math.round(Math.max(0, Math.min(1, value)) * 255).toString(16).padStart(2, '0').toUpperCase()
  return `#${byte(r)}${byte(g)}${byte(b)}`
}

/**
 * One ExcelJS cell style (`cell.style`) to the grid's CSS shape:
 * `{ font, fill, border, align }`, where `fill` is `#RRGGBB` or `null`.
 * `notes` collects the parts that could not be resolved, for the caller to surface once.
 */
export function resolveStyle(style, palette) {
  const notes = []
  const source = style ?? {}
  return {
    font: resolveFont(source.font, palette, notes),
    fill: resolveFill(source.fill, palette, notes),
    border: resolveBorder(source.border, palette, notes),
    align: resolveAlign(source.alignment),
    notes
  }
}

function resolveFont(font, palette, notes) {
  if (!font) return { bold: false, italic: false, underline: null, strike: false, size: null, color: null, family: null, vertAlign: null }
  return {
    bold: font.bold === true,
    italic: font.italic === true,
    underline: normalizeUnderline(font.underline),
    strike: font.strike === true,
    size: typeof font.size === 'number' ? font.size : null,
    color: cssColor(font.color, palette, notes),
    // `family` is the CSS font-family name; ExcelJS's numeric `family` is a font classification.
    family: typeof font.name === 'string' ? font.name : null,
    vertAlign: /^(superscript|subscript)$/.test(font.vertAlign) ? font.vertAlign : null
  }
}

function normalizeUnderline(underline) {
  if (underline === undefined || underline === null || underline === false) return null
  if (underline === 'double') return 'double'
  return 'single'
}

function resolveFill(fill, palette, notes) {
  if (!fill || fill.type !== 'pattern') return null
  const pattern = fill.pattern
  if (!pattern || pattern === 'none') return null
  const color = cssColor(fill.fgColor, palette, notes)
  if (!color) {
    if (fill.fgColor && pattern !== 'solid') notes.push(`pattern fill "${pattern}" has no resolvable foreground colour`)
    return null
  }
  if (pattern !== 'solid') notes.push(`pattern fill "${pattern}" is drawn as a solid ${color}`)
  return color
}

function resolveBorder(border, palette, notes) {
  return {
    top: borderSide(border?.top, palette, notes),
    right: borderSide(border?.right, palette, notes),
    bottom: borderSide(border?.bottom, palette, notes),
    left: borderSide(border?.left, palette, notes)
  }
}

function borderSide(side, palette, notes) {
  if (!side || !side.style) return null
  return { style: side.style, color: cssColor(side.color, palette, notes) }
}

function resolveAlign(alignment) {
  if (!alignment) return { horizontal: null, vertical: null, wrapText: false, indent: 0, textRotation: 0 }
  return {
    horizontal: alignment.horizontal ?? null,
    // ExcelJS spells a vertically centred cell "middle"; CSS spells it "center".
    vertical: alignment.vertical === 'middle' ? 'center' : alignment.vertical ?? null,
    wrapText: alignment.wrapText === true,
    indent: typeof alignment.indent === 'number' ? alignment.indent : 0,
    textRotation: normalizeRotation(alignment.textRotation)
  }
}

/** OOXML stores rotations of -90..-1 degrees as 91..180. */
function normalizeRotation(rotation) {
  if (typeof rotation !== 'number' || rotation === 0) return 0
  if (rotation > 90 && rotation <= 180) return 90 - rotation
  return rotation
}
