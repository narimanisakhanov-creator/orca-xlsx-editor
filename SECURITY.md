# Security

## What this plugin can do

XLSX Editor runs in two parts, and they hold different power.

**The editor frame** (`sandboxed-ui`) draws a workbook. It receives the file's bytes from Orca and
hands back new bytes. It never learns the file's path, has no file system, no network, and no
clipboard API. It is a page inside `orca-plugin-frame://` with `default-src 'self'`.

**The worker** (`trusted-worker`, `worker/main.js`) is a Node process. It exists for one reason: the
**В агента** hand-off, which needs two host APIs the frame cannot call.

## The two capabilities

| Capability | What it is used for | Exact call |
| --- | --- | --- |
| `agent:read` | listing the agents of the focused workspace, so the hand-off can offer a target | `agents.list`, nothing else |
| `terminal:send` | typing the chosen text into the chosen terminal | `terminal.sendText({ terminalId, text, enter })`, one call per accepted request |

Both are used only from `worker/main.js`, only inside its two panel handlers, and only after a press.

## What it never does

- Never reads, writes, or deletes a file. The workbook reaches it as bytes; the path is Orca's.
- Never opens a network connection. The frame's CSP refuses requests outside the plugin's own files.
- Never constructs code from strings: the build fails if `eval` or a `Function` constructor appears in
  anything it emits.
- Never sends anything on its own. There is no timer, no hook, no event subscription; the worker
  starts when the frame asks and goes idle when it finishes.
- Never writes, or saves, a normalized workbook. The parser-only copy built for ExcelJS exists in
  memory; a save starts from the accepted file's own ZIP entries.
- Never touches a terminal the person did not pick from the list the plugin itself showed.

## What the hand-off sends

`agent.send` takes `{ terminalId, text, enter? }`. It is refused before any host call when the
`terminalId` is empty or too long, `text` is missing or empty, `enter` is not a boolean, or `text`
is longer than 4096 characters. That limit is Orca's on `terminal.sendText`; the hand-off is never
split or truncated, because Orca offers no atomicity or idempotency across calls. For text over the
message limit, the popup offers a button to copy the full hand-off for manual pasting. A selection
over 200,000 cells exceeds the frame's separate limit: both sending and copying are refused. Before
materializing the copy, a conservative estimate including a header allowance is limited to
8 × 1024 × 1024 UTF-16 code units; it may refuse a smaller actual output. If either copy limit is
exceeded, nothing is copied and the person is asked to reduce the selection.

A valid request makes exactly one `terminal.sendText` call, with `enter: true` only when the person
asked for it. One send per terminal is allowed at a time: a second request while the first is in
flight is refused. The answer is `{ accepted, sent, unknown, reason? }`. On success `sent` is the
message the host accepted. A non-accepting host result is a refusal with `unknown: false`. A thrown
host call returns `unknown: true` unless its code proves the host refused before typing
(`invalid_request`, `invalid_params`, `unknown_method`, `capability_denied`, `consent_required`,
`not_permitted`, `rate_limited`, `too_large`); after an unknown outcome the person must inspect the
terminal, and nothing is retried automatically.

`enter` defaults to **false**, so the text lands in the input line for the person to read and send
themselves. When "Send immediately (Enter)" is ticked, that one call carries `enter: true`. For a
sendable selection, the preview in the popup is exactly the text that call will contain.

The text is a table: a one-line context header (sheet name and range) followed by the selection as a
Markdown table or TSV. No cell outside the selection, no formula the person did not ask for, and no
workbook metadata is included.

## The reading copy

Some workbooks need a parser-only compatibility pass before ExcelJS can read them. That pass builds
a copy in memory: XML namespace prefixes and their declarations are stripped, `xl/drawings/*`
(except `.vml`) and `xl/tables/*` parts are removed, `<drawing>` and `<tableParts>` references are
removed from the sheet XML, and sheet relationship targets are normalized to paths relative to the
sheet (absolute targets and `.`/`..` segments included). The copy is never written to a file and is
never the base of a save — saving starts from the accepted file's own ZIP entries. When a sheet
carries conditional formatting, the editor shows a visible warning that the rules are preserved but
not displayed. An unreadable archive, or a workbook where ExcelJS returns a different number of
sheets than `xl/workbook.xml` declares, is an error rather than a partial workbook.

## Structural edits

Inserting or deleting rows or columns rewrites the original package entries atomically: every sheet,
defined name, merged range, table, drawing, filter, selection, validation, hyperlink and column
definition is translated before any entry is replaced. Unsupported formulas (dynamic, 3D or external
references), protected sheets, extensions, charts and unresolved object relationships refuse the
whole operation; no partial package is ever recorded. The history keeps the before/after entry
snapshots in memory so undo and redo work even after the host accepted a save. Macro projects
(`xl/vbaProject.bin`) are opaque and copied byte-for-byte; VBA code is not rewritten. Structural
history is lost when the document is reloaded or replaced, as is ordinary undo.

## Reporting a problem

Open an issue at https://github.com/narimanisakhanov-creator/orca-xlsx-editor/issues, or write
privately through GitHub's security advisory form on that repository:
https://github.com/narimanisakhanov-creator/orca-xlsx-editor/security/advisories/new