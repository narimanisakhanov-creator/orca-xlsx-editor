// Builds the workbook fixtures the tests read, so no spreadsheet binary has to be committed.
//
// The parts are plain OOXML strings and fflate packs them. That is deliberate: the code under test
// is our reader (`src/model/workbook.js` on top of ExcelJS) and our writer (`src/write/zip.js`), and
// a fixture produced by either of them would hide the bugs the tests are meant to catch. fflate is
// a third-party codec, not a spreadsheet implementation.
//
// Every zip entry carries the same fixed DOS timestamp, so two runs produce byte-identical books.
// The books themselves are written to `test/fixtures/`, which is git-ignored and may be deleted at
// any time — this generator is the only thing that recreates them.
//
//   npm run fixtures      (or: node test/helpers/build-fixtures.mjs)
//
// `test/helpers/fixtures.mjs` is what the specs import: it builds in memory on first use, so
// `git clone && npm ci && npm test` works without any preparation and without Python.
import { mkdirSync, writeFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

import { zipSync } from 'fflate'

const HERE = path.dirname(fileURLToPath(import.meta.url))

/** Where the built books land; created on demand and git-ignored. */
export const OUTPUT_DIR = path.join(HERE, '..', 'fixtures')

/** A zip record carries a DOS timestamp; a real date would make every run differ. */
const FIXED_MTIME = new Date(Date.UTC(2020, 0, 1, 0, 0, 0))
const DEFLATE_LEVEL = 6

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const THEME_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

const HEAD_FILL = 'FF1F4E79'
const BAND_FILL = 'FFF2F5F8'
const BORDER_COLOR = 'FFB0B0B0'
const CONDITIONAL_FILL = 'FFFFC7CE'

/* ---------------------------------------------------------------- shared parts ---------------- */

/** Office theme, trimmed to the elements Excel and this plugin read: the colour scheme and fonts. */
const THEME = `${XML}<a:theme xmlns:a="${THEME_NS}" name="Office Theme"><a:themeElements>` +
  '<a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>' +
  '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>' +
  '<a:accent6><a:srgbClr val="F79646"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>' +
  '</a:clrScheme>' +
  '<a:fontScheme name="Office">' +
  '<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
  '</a:fontScheme>' +
  '<a:fmtScheme name="Office">' +
  '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
  '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
  '<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
  '<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
  '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'

const APP_PROPS = `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
  '<Application>orca-xlsx-editor test fixtures</Application><AppVersion>1.0</AppVersion></Properties>'

const CORE_PROPS = `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
  '<dc:creator>orca-xlsx-editor tests</dc:creator>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified>' +
  '</cp:coreProperties>'

const ROOT_RELS = `${XML}<Relationships xmlns="${PKG_REL_NS}">` +
  `<Relationship Type="${REL_NS}/officeDocument" Target="xl/workbook.xml" Id="rId1" />` +
  `<Relationship Type="${PKG_REL_NS}/metadata/core-properties" Target="docProps/core.xml" Id="rId2" />` +
  `<Relationship Type="${REL_NS}/extended-properties" Target="docProps/app.xml" Id="rId3" />` +
  '</Relationships>'

const STYLE_PART = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const THEME_PART = 'application/vnd.openxmlformats-officedocument.theme+xml'
const CORE_PART = 'application/vnd.openxmlformats-package.core-properties+xml'
const APP_PART = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
const SHEET_PART = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
const WORKBOOK_PART = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
const VBA_PART = 'application/vnd.ms-office.vbaProject'

/** `[Content_Types].xml` for a package whose overrides are given in the order they should appear. */
function contentTypes(overrides) {
  const defaults =
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' +
    '<Default Extension="xml" ContentType="application/xml" />'
  const body = overrides
    .map(([partName, contentType]) => `<Override PartName="${partName}" ContentType="${contentType}" />`)
    .join('')
  return `${XML}<Types xmlns="${TYPES_NS}">${defaults}${body}</Types>`
}

/* ---------------------------------------------------------------- cell helpers ---------------- */

/** `<c>` text: only the XML metacharacters can appear in the values these books carry. */
function escape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/** 1 -> A, 27 -> AA. */
function columnName(index) {
  let name = ''
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name
  }
  return name
}

const styleOf = (style) => (style === undefined ? '' : ` s="${style}"`)

const textCell = (ref, value, style) =>
  `<c r="${ref}"${styleOf(style)} t="inlineStr"><is><t>${escape(value)}</t></is></c>`

const numberCell = (ref, value, style) =>
  `<c r="${ref}"${styleOf(style)} t="n"><v>${escape(value)}</v></c>`

/**
 * A formula with no cached result, which is what a writer that does not evaluate produces: the
 * model then falls back to the formula text, and the writer must keep the empty `<v />` untouched.
 */
const formulaCell = (ref, formula, style) =>
  `<c r="${ref}"${styleOf(style)}><f>${escape(formula)}</f><v /></c>`

const colElement = (index, width) =>
  `<col width="${width}" customWidth="1" min="${index}" max="${index}" />`

/* ---------------------------------------------------------------- styles ---------------------- */

/**
 * The style table of `rich.xlsx`. The indexes are part of the fixture's contract: the writer tests
 * assert that an edit keeps `s="3"` on B2 or `s="15"` on A9.
 *
 *   0 default                  8 banded '#,##0', right aligned
 *   1 header                   9 banded currency
 *   2 bordered                10 banded percent
 *   3 '#,##0', right           11 banded date
 *   4 currency                12 bold
 *   5 percent                 13 bold '#,##0'
 *   6 date                    14 '#,##0.00'
 *   7 banded bordered         15 centered (merged note)
 */
const RICH_STYLES = `${XML}<styleSheet xmlns="${NS}">` +
  '<numFmts count="2">' +
  '<numFmt numFmtId="164" formatCode="#,##0.00 &quot;RUB&quot;" />' +
  '<numFmt numFmtId="165" formatCode="yyyy-mm-dd" />' +
  '</numFmts>' +
  '<fonts count="3">' +
  '<font><name val="Calibri" /><family val="2" /><color theme="1" /><sz val="11" /><scheme val="minor" /></font>' +
  '<font><b val="1" /><color rgb="FFFFFFFF" /><sz val="11" /></font>' +
  '<font><b val="1" /></font>' +
  '</fonts>' +
  '<fills count="4">' +
  '<fill><patternFill /></fill>' +
  '<fill><patternFill patternType="gray125" /></fill>' +
  `<fill><patternFill patternType="solid"><fgColor rgb="${HEAD_FILL}" /></patternFill></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="${BAND_FILL}" /></patternFill></fill>` +
  '</fills>' +
  '<borders count="2">' +
  '<border><left /><right /><top /><bottom /><diagonal /></border>' +
  `<border><left style="thin"><color rgb="${BORDER_COLOR}" /></left>` +
  `<right style="thin"><color rgb="${BORDER_COLOR}" /></right>` +
  `<top style="thin"><color rgb="${BORDER_COLOR}" /></top>` +
  `<bottom style="thin"><color rgb="${BORDER_COLOR}" /></bottom></border>` +
  '</borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>' +
  '<cellXfs count="16">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">' +
  '<alignment horizontal="center" vertical="center" wrapText="1" /></xf>' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">' +
  '<alignment horizontal="right" /></xf>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="10" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="3" fontId="0" fillId="3" borderId="1" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">' +
  '<alignment horizontal="right" /></xf>' +
  '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="10" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="165" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="3" fontId="2" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">' +
  '<alignment horizontal="center" /></xf>' +
  '</cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0" hidden="0" /></cellStyles>' +
  `<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="${CONDITIONAL_FILL}" />` +
  '</patternFill></fill></dxf></dxfs>' +
  '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />' +
  '</styleSheet>'

/**
 * The style table of `styled.xlsx`: the banded variant first so `A2` is `s="2"` and the same
 * bordered-but-unpainted block after it so an odd row starts at `s="8"`.
 *
 *   0 default                 7 banded '#,##0.0'
 *   1 header                  8 bordered
 *   2 banded bordered         9 bordered '#,##0.00'
 *   3 banded '#,##0.00'      10 bordered '#,##0'
 *   4 banded '#,##0'         11 bordered '[Red]-#,##0'
 *   5 banded '[Red]-#,##0'   12 bordered percent
 *   6 banded percent         13 bordered '#,##0.0'
 */
const STYLED_STYLES = `${XML}<styleSheet xmlns="${NS}">` +
  '<numFmts count="2">' +
  '<numFmt numFmtId="164" formatCode="#,##0;[Red]-#,##0" />' +
  '<numFmt numFmtId="165" formatCode="#,##0.0" />' +
  '</numFmts>' +
  '<fonts count="2">' +
  '<font><name val="Calibri" /><family val="2" /><color theme="1" /><sz val="11" /><scheme val="minor" /></font>' +
  '<font><b val="1" /><color rgb="FFFFFFFF" /><sz val="11" /></font>' +
  '</fonts>' +
  '<fills count="4">' +
  '<fill><patternFill /></fill>' +
  '<fill><patternFill patternType="gray125" /></fill>' +
  `<fill><patternFill patternType="solid"><fgColor rgb="${HEAD_FILL}" /></patternFill></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="${BAND_FILL}" /></patternFill></fill>` +
  '</fills>' +
  '<borders count="2">' +
  '<border><left /><right /><top /><bottom /><diagonal /></border>' +
  `<border><left style="thin"><color rgb="${BORDER_COLOR}" /></left>` +
  `<right style="thin"><color rgb="${BORDER_COLOR}" /></right>` +
  `<top style="thin"><color rgb="${BORDER_COLOR}" /></top>` +
  `<bottom style="thin"><color rgb="${BORDER_COLOR}" /></bottom></border>` +
  '</borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>' +
  '<cellXfs count="14">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyAlignment="1" pivotButton="0" quotePrefix="0" xfId="0">' +
  '<alignment horizontal="center" vertical="center" wrapText="1" /></xf>' +
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="4" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="3" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="10" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="165" fontId="0" fillId="3" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="10" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" pivotButton="0" quotePrefix="0" xfId="0" />' +
  '</cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0" hidden="0" /></cellStyles>' +
  '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />' +
  '</styleSheet>'

/* ---------------------------------------------------------------- rich.xlsx: Summary ---------- */

const SUMMARY_HEADER = ['Store', 'Sales', 'Revenue', 'Margin', 'Opened']
const SUMMARY_WIDTHS = [44, 13, 20, 10, 14]
// Store, sales, margin, opened — revenue is the formula column and is written from the row number.
// The date column holds text on purpose: the spreadsheet the fixture stands in for stored it so.
const SUMMARY_ROWS = [
  ['5063GS_Krasn_1', 282349, 0.0812, '2019-04-01'],
  ['4047M_Lazorevy', 271219, 0.064, '2020-11-15'],
  ['4835DS_Lefortovo', 262386, 0.03, '2018-06-30'],
  ['5005DS_Kuzovlevo', 242698, 0.05, '2021-02-12'],
  ['5818DS_Olshansk', 237344, 0.06, '2017-09-08']
]

const summaryRow = (index) => {
  const r = index + 2
  // Odd rows carry the banded fill; the style indexes follow RICH_STYLES, where the banded block
  // is five slots after the plain one.
  const base = r % 2 === 1 ? 7 : 2
  const [store, sales, margin, opened] = SUMMARY_ROWS[index]
  return `<row r="${r}">` +
    textCell(`A${r}`, store, base) +
    numberCell(`B${r}`, sales, base + 1) +
    formulaCell(`C${r}`, `B${r}*48.7`, base + 2) +
    numberCell(`D${r}`, margin, base + 3) +
    textCell(`E${r}`, opened, base + 4) +
    '</row>'
}

const SUMMARY_SHEET = `${XML}<worksheet xmlns="${NS}">` +
  '<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>' +
  '<dimension ref="A1:E9" />' +
  '<sheetViews><sheetView workbookViewId="0">' +
  '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen" />' +
  '<selection pane="topRight" /><selection pane="bottomLeft" />' +
  '<selection pane="bottomRight" activeCell="A1" sqref="A1" />' +
  '</sheetView></sheetViews>' +
  '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />' +
  `<cols>${SUMMARY_WIDTHS.map((width, index) => colElement(index + 1, width)).join('')}</cols>` +
  '<sheetData>' +
  `<row r="1" ht="30" customHeight="1">${SUMMARY_HEADER.map((name, index) =>
    textCell(`${columnName(index + 1)}1`, name, 1)).join('')}</row>` +
  SUMMARY_ROWS.map((unused, index) => summaryRow(index)).join('') +
  '<row r="7">' +
  textCell('A7', 'Total', 12) +
  formulaCell('B7', 'SUM(B2:B6)', 13) +
  formulaCell('C7', 'SUM(C2:C6)', 14) +
  '</row>' +
  '<row r="9">' + textCell('A9', 'Prepared from the store ledger', 15) + '</row>' +
  '</sheetData>' +
  '<mergeCells count="1"><mergeCell ref="A9:E9" /></mergeCells>' +
  '<conditionalFormatting sqref="D2:D6">' +
  '<cfRule type="cellIs" priority="1" operator="lessThan" dxfId="0"><formula>0.04</formula></cfRule>' +
  '</conditionalFormatting>' +
  '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" />' +
  '</worksheet>'

/* ---------------------------------------------------------------- rich.xlsx: Notes / Plain ----- */

const NOTES_SHEET = `${XML}<worksheet xmlns="${NS}">` +
  '<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>' +
  '<dimension ref="A1:A3" />' +
  '<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1" /></sheetView></sheetViews>' +
  '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />' +
  '<cols><col width="18" customWidth="1" min="1" max="1" /></cols>' +
  '<sheetData>' +
  '<row r="1">' + textCell('A1', 'Sheet with no styling') + '</row>' +
  '<row r="3">' + textCell('A3', 'A longer line of text that should stay readable in a narrow column') + '</row>' +
  '</sheetData>' +
  '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" />' +
  '</worksheet>'

const PLAIN_ROWS = 200

const PLAIN_SHEET = `${XML}<worksheet xmlns="${NS}">` +
  '<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>' +
  `<dimension ref="A1:B${PLAIN_ROWS}" />` +
  '<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1" /></sheetView></sheetViews>' +
  '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />' +
  '<sheetData>' +
  Array.from({ length: PLAIN_ROWS }, (unused, index) => {
    const r = index + 1
    return `<row r="${r}">${numberCell(`A${r}`, r)}${textCell(`B${r}`, `row ${r}`)}</row>`
  }).join('') +
  '</sheetData>' +
  '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" />' +
  '</worksheet>'

/* ---------------------------------------------------------------- styled.xlsx: Ledger ---------- */

const LEDGER_ROWS = 166
const LEDGER_COLUMNS = 16
const LEDGER_HEADER = ['Store', 'Sales', 'Revenue', 'Plan', 'Delta', 'Margin', 'Staff', 'Area']
  .concat(Array.from({ length: 8 }, (unused, index) => `W${index + 1}`))

/**
 * Style indexes of one Ledger body row. `STYLED_STYLES` lists the banded block (2..7) first and
 * the same shapes on a white background after it (8..13), so one table serves both kinds of row.
 */
const rowStyles = (banded) => ({
  text: banded ? 2 : 8,
  money: banded ? 3 : 9,
  integer: banded ? 4 : 10,
  red: banded ? 5 : 11,
  percent: banded ? 6 : 12,
  decimal: banded ? 7 : 13
})

/** One Ledger body row: the formats the tests read live in STYLED_STYLES, not in the sheet. */
const ledgerRow = (r) => {
  // Even rows carry the banded fill and start at style 2; odd ones are plain and start at style 8.
  const style = rowStyles(r % 2 === 0)
  const store = `${4000 + r}_${`Store ${r}`.padEnd(20, '0')}`.slice(0, 40)
  const cells = [
    textCell(`A${r}`, store, style.text),
    numberCell(`B${r}`, 200000 + r * 137, style.money),
    formulaCell(`C${r}`, `B${r}*48.7`, style.money),
    numberCell(`D${r}`, 250000 + r * 100, style.integer),
    formulaCell(`E${r}`, `B${r}-D${r}`, style.red),
    numberCell(`F${r}`, (r % 15) / 100, style.percent),
    numberCell(`G${r}`, 5 + (r % 7), style.text),
    numberCell(`H${r}`, 80 + (r % 40), style.decimal)
  ]
  for (let index = 9; index <= LEDGER_COLUMNS; index++) {
    cells.push(numberCell(`${columnName(index)}${r}`, (r * index) % 997, style.integer))
  }
  return `<row r="${r}">${cells.join('')}</row>`
}

const LEDGER_SHEET = `${XML}<worksheet xmlns="${NS}">` +
  '<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>' +
  `<dimension ref="A1:P${LEDGER_ROWS}" />` +
  '<sheetViews><sheetView workbookViewId="0">' +
  '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen" />' +
  '<selection pane="topRight" /><selection pane="bottomLeft" />' +
  '<selection pane="bottomRight" activeCell="A1" sqref="A1" />' +
  '</sheetView></sheetViews>' +
  '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />' +
  '<cols>' +
  Array.from({ length: LEDGER_COLUMNS }, (unused, index) => colElement(index + 1, index === 0 ? 43 : 14)).join('') +
  '</cols>' +
  '<sheetData>' +
  `<row r="1" ht="30" customHeight="1">${LEDGER_HEADER.map((name, index) =>
    textCell(`${columnName(index + 1)}1`, name, 1)).join('')}</row>` +
  Array.from({ length: LEDGER_ROWS - 1 }, (unused, index) => ledgerRow(index + 2)).join('') +
  '</sheetData>' +
  '<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" />' +
  '</worksheet>'

/* ---------------------------------------------------------------- workbook parts --------------- */

const WORKSHEET_REL = `${REL_NS}/worksheet`

/** `xl/workbook.xml`; the sheets are addressed through the relationship ids given here. */
function workbookPart(names) {
  const sheets = names
    .map((name, index) => `<sheet name="${name}" sheetId="${index + 1}" state="visible" r:id="rId${index + 1}" />`)
    .join('')
  const definedNames = names.includes('Summary')
    ? '<definedNames><definedName name="SalesTotal">Summary!$B$7</definedName></definedNames>'
    : ''
  return `${XML}<workbook xmlns:r="${REL_NS}" xmlns="${NS}">` +
    '<workbookPr />' +
    '<bookViews><workbookView visibility="visible" minimized="0" showHorizontalScroll="1" ' +
    'showVerticalScroll="1" showSheetTabs="1" tabRatio="600" firstSheet="0" activeTab="0" ' +
    'autoFilterDateGrouping="1" /></bookViews>' +
    `<sheets>${sheets}</sheets>` +
    definedNames +
    // fullCalcOnLoad is what the writer must not have to add: the fixtures ship a book that already
    // asks Excel to recalculate, and the writer tests rely on that to keep an edit to one part.
    '<calcPr calcId="124519" fullCalcOnLoad="1" />' +
    '</workbook>'
}

function workbookRels(count) {
  const sheets = Array.from({ length: count }, (unused, index) =>
    `<Relationship Type="${WORKSHEET_REL}" Target="/xl/worksheets/sheet${index + 1}.xml" Id="rId${index + 1}" />`)
    .join('')
  return `${XML}<Relationships xmlns="${PKG_REL_NS}">` +
    sheets +
    `<Relationship Type="${REL_NS}/styles" Target="styles.xml" Id="rId${count + 1}" />` +
    `<Relationship Type="${REL_NS}/theme" Target="theme/theme1.xml" Id="rId${count + 2}" />` +
    '</Relationships>'
}

function worksheetContentTypes(count) {
  return Array.from({ length: count }, (unused, index) => [`/xl/worksheets/sheet${index + 1}.xml`, SHEET_PART])
}

function packageParts({ sheets, styles, workbook, rels, types }) {
  return [
    { name: 'docProps/app.xml', data: APP_PROPS },
    { name: 'docProps/core.xml', data: CORE_PROPS },
    { name: 'xl/theme/theme1.xml', data: THEME },
    ...sheets,
    { name: 'xl/styles.xml', data: styles },
    { name: '_rels/.rels', data: ROOT_RELS },
    { name: 'xl/workbook.xml', data: workbook },
    { name: 'xl/_rels/workbook.xml.rels', data: rels },
    { name: '[Content_Types].xml', data: types }
  ]
}

/* ---------------------------------------------------------------- the three books -------------- */

function richParts() {
  const sheets = [
    { name: 'xl/worksheets/sheet1.xml', data: SUMMARY_SHEET },
    { name: 'xl/worksheets/sheet2.xml', data: NOTES_SHEET },
    { name: 'xl/worksheets/sheet3.xml', data: PLAIN_SHEET }
  ]
  const types = contentTypes([
    ['/xl/styles.xml', STYLE_PART],
    ['/xl/theme/theme1.xml', THEME_PART],
    ['/docProps/core.xml', CORE_PART],
    ['/docProps/app.xml', APP_PART],
    ...worksheetContentTypes(3),
    ['/xl/workbook.xml', WORKBOOK_PART]
  ])
  return packageParts({
    sheets,
    styles: RICH_STYLES,
    workbook: workbookPart(['Summary', 'Notes', 'Plain']),
    rels: workbookRels(3),
    types
  })
}

/** Deterministic stand-in for a VBA project: 10 240 bytes, the same on every run. */
function vbaProject() {
  const bytes = new Uint8Array(10240)
  for (let index = 0; index < bytes.length; index++) bytes[index] = index % 256
  return bytes
}

function macroParts() {
  const parts = richParts().map((part) =>
    part.name === '[Content_Types].xml'
      ? {
          name: part.name,
          data: part.data.replace(
            '</Types>',
            `<Override PartName="/xl/vbaProject.bin" ContentType="${VBA_PART}"/></Types>`
          )
        }
      : part)
  parts.push({ name: 'xl/vbaProject.bin', data: vbaProject() })
  return parts
}

function styledParts() {
  const sheets = [{ name: 'xl/worksheets/sheet1.xml', data: LEDGER_SHEET }]
  const types = contentTypes([
    ['/xl/styles.xml', STYLE_PART],
    ['/xl/theme/theme1.xml', THEME_PART],
    ['/docProps/core.xml', CORE_PART],
    ['/docProps/app.xml', APP_PART],
    ...worksheetContentTypes(1),
    ['/xl/workbook.xml', WORKBOOK_PART]
  ])
  return packageParts({
    sheets,
    styles: STYLED_STYLES,
    workbook: workbookPart(['Ledger']),
    rels: workbookRels(1),
    types
  })
}

/** Packs parts into a deterministic archive: same names, same order, same timestamp, same bytes. */
function pack(parts) {
  const encoder = new TextEncoder()
  const files = {}
  for (const part of parts) {
    files[part.name] = part.data instanceof Uint8Array ? part.data : encoder.encode(part.data)
  }
  return zipSync(files, { level: DEFLATE_LEVEL, mtime: FIXED_MTIME })
}

/**
 * The three fixtures, in memory. Pure: no filesystem, no clock, no randomness — two calls in the
 * same process, or two runs of `npm test`, produce identical bytes.
 *
 * @returns {{ 'rich.xlsx': Uint8Array, 'macro.xlsm': Uint8Array, 'styled.xlsx': Uint8Array }}
 */
export function buildFixtures() {
  return {
    'rich.xlsx': pack(richParts()),
    'macro.xlsm': pack(macroParts()),
    'styled.xlsx': pack(styledParts())
  }
}

/** Builds the books and writes them into `directory` (default `test/fixtures/`); returns their paths. */
export function writeFixtures(directory = OUTPUT_DIR) {
  mkdirSync(directory, { recursive: true })
  return Object.entries(buildFixtures()).map(([name, bytes]) => {
    const target = path.join(directory, name)
    writeFileSync(target, bytes)
    console.log(`wrote ${target} (${bytes.byteLength} bytes)`)
    return target
  })
}

// `node test/helpers/build-fixtures.mjs` (or `npm run fixtures`): regenerate the books by hand.
if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  writeFixtures()
}
