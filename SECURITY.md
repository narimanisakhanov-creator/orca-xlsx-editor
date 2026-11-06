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
| `terminal:send` | typing the chosen text into the chosen terminal | `terminal.sendText({ terminalId, text, enter })` |

Both are used only from `worker/main.js`, only inside the two handlers below, and only after a press.

## What it never does

- Never reads, writes, or deletes a file. The workbook reaches it as bytes; the path is Orca's.
- Never opens a network connection. The frame's CSP refuses requests outside the plugin's own files.
- Never constructs code from strings: the build fails if `eval` or a `Function` constructor appears in
  anything it emits.
- Never sends anything on its own. There is no timer, no hook, no event subscription; the worker
  starts when the frame asks and goes idle when it finishes.
- Never touches a terminal the person did not pick from the list the plugin itself showed.

## What the hand-off sends

`agent.send` takes `{ terminalId, text, enter }`, refuses an empty `terminalId` or an empty `text`,
cuts `text` to Orca's 4096-character limit on `terminal.sendText` (on a line boundary) and returns
`{ accepted, sent, truncated }`. `enter` defaults to **false**, so the text lands in the input line
for the person to read and send themselves. The preview in the popup is produced by the same
function, so what is shown is what is typed.

The text is a table: a one-line context header (sheet name and range) followed by the selection as a
Markdown table or TSV. No cell outside the selection, no formula the person did not ask for, and no
workbook metadata is included.

## Reporting a problem

Open an issue at https://github.com/narimanisakhanov-creator/orca-xlsx-editor/issues, or write
privately through GitHub's security advisory form on that repository:
https://github.com/narimanisakhanov-creator/orca-xlsx-editor/security/advisories/new