# Troubleshooting

## First step: run `excel_check_environment`

In any chat with Claude, ask:

> Run `excel_check_environment` and tell me what it says.

The response is a structured capability report. The `recommendations` array tells you exactly what to fix.

## Common failures

### `PATH_OUTSIDE_ALLOWED`

**What:** You tried to read or write a file outside the configured sandbox folders. By default the server only allows your `Documents`, `Desktop`, and `Downloads` folders.

**Fix:** Add the folder you want to access via the `Allowed Directories` setting (Claude Desktop Extensions UI) or `EXCEL_ALLOWED_DIRS` env var (manual config).

### Tool hangs forever on Mac

**What:** macOS Automation permission has not been granted yet. The first `osascript` call from the server triggers a permission dialog that may be hidden behind another window.

**Fix:**
1. Switch focus to **Microsoft Excel** so it's frontmost.
2. Switch back to Claude — the prompt should now be visible.
3. Click **OK**.
4. (Alternative) System Settings → **Privacy & Security** → **Automation** → expand the entry for Claude Desktop / Claude Code → enable **Microsoft Excel**.

This is one-time per Claude install.

### `VBA_TRUST_ERROR` / "VBA access denied" on Windows

**What:** Excel's Trust Center blocks programmatic access to the VBA project object model.

**Fix:** Excel → **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings** → tick **Trust access to the VBA project object model**. Restart Excel.

### "PowerShell timed out" on Windows

**What:** Excel has a modal dialog open (Save prompt, VBA error popup, "Format Cells" dialog left open from a click, etc.). PowerShell's COM call cannot proceed while a modal is up.

**Fix:** Switch to Excel, dismiss any open dialog, retry the tool. If it persists, run `excel_diagnose_connection`.

### Tool says "Requires Windows" but you're on Mac

**What:** Some live-mode tools (VBA macros, native chart/pivot creation while Excel is open) only work via Windows COM. Mac doesn't have an equivalent OS-level automation path for those features.

**Fix:** Most have a file-mode equivalent that works on Mac. The error message tells you which one. See [PLATFORM_PARITY.md](PLATFORM_PARITY.md) for the full matrix.

### Excel for the Web (no desktop Excel)

**What:** Live editing tools require the desktop Excel app — they talk to it via COM (Windows) or AppleScript (Mac). The browser version doesn't expose those APIs.

**What still works:** All file-mode tools (~38 tools): open, read, edit, format, chart, find/replace, CSV import/export, save. The server edits the `.xlsx` directly. You won't see live changes, but the file is updated correctly.

**What doesn't work:** Live tools that require Excel running (~24 tools — VBA, real-time formatting, screenshot capture, Power Query refresh).

### Claude Code: server not loading

Check `~/.claude.json` for valid JSON. A missing comma in the `mcpServers` block silently disables ALL servers in that project. Quick check:

```bash
node -e "JSON.parse(require('fs').readFileSync(process.env.USERPROFILE + '/.claude.json'))"
```

(or `process.env.HOME` on Mac/Linux)

If it errors, your file has a JSON syntax problem at the line/column shown.

### Excel won't reopen the file after editing

If the server saved the file while you had it open in Excel, Excel may show a "file changed externally" dialog. Click **Discard Changes** (your file-mode edit is the canonical version) or close Excel without saving and reopen.

To avoid this: close the file in Excel before invoking file-mode write tools, OR use the live-mode tools (which talk to Excel directly while it's open).

## Getting help

If `excel_check_environment` shows everything green and you're still stuck, capture:

1. The exact prompt you sent.
2. The full Claude response (including the JSON tool call and tool result).
3. Output of `excel_check_environment`.

…and ask in the support channel.

## The `__excel_mcp_idempotency__` sheet

If you see a sheet named `__excel_mcp_idempotency__` in your workbook, that's an internal bookkeeping sheet used by the dedupKey feature. It's marked very-hidden so you won't see it normally. Don't edit it — the server uses it to detect duplicate operations.
