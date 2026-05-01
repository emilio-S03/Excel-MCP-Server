# Install — Excel MCP Server v3.0.0

Two ways to install. Pick one.

## Option A: One-click (Claude Desktop)

1. Download `excel-mcp-server-3.0.0.mcpb` (the file ends in `.mcpb`, not `.zip` or `Source code`).
2. Double-click it.
3. Claude Desktop opens an installer dialog → click **Install**.
4. **macOS only:** if you see "Apple could not verify… is free of malware", right-click the file → **Open** → **Open**. One-time.
5. **Windows only:** if you see SmartScreen "Windows protected your PC", click **More info** → **Run anyway**. One-time.
6. Restart Claude Desktop fully (right-click the tray icon → **Quit**, then reopen — closing the window is not enough).
7. Open a chat. Click the tools icon at the bottom — you should see 62 Excel tools listed.

After install, in Claude Desktop **Settings → Extensions → Excel MCP Server** you can configure:

- **Allowed Directories** — folders this server is allowed to read/write. Defaults to your `Documents`, `Desktop`, and `Downloads`. Add your project folder here if you keep `.xlsx` files outside those.
- **Create Backup by Default** — leave **off** unless you specifically want a `.backup` copy made before every modification.
- **Default Response Format** — leave **json**.

## Option B: Manual install (Claude Desktop or Claude Code)

Edit your config file:

- **Claude Desktop:** `%APPDATA%\Claude\claude_desktop_config.json` (Windows) or `~/Library/Application Support/Claude/claude_desktop_config.json` (Mac)
- **Claude Code:** `~/.claude.json` (look for the `mcpServers` block in your project entry)

Add inside the `mcpServers` block (Windows path example):

```json
"excel": {
  "command": "node",
  "args": ["C:\\path\\to\\excel-mcp-server\\dist\\index.js"],
  "env": {
    "EXCEL_ALLOWED_DIRS": "C:/Users/YOU/Documents;C:/Users/YOU/Desktop;C:/Users/YOU/Downloads;C:/Users/YOU/Projects",
    "EXCEL_CREATE_BACKUP_BY_DEFAULT": "false",
    "EXCEL_DEFAULT_RESPONSE_FORMAT": "json"
  }
}
```

On macOS, separate `EXCEL_ALLOWED_DIRS` paths with `:` instead of `;`, and use `/Users/YOU/...` paths.

Then: build the server (`npm install && npm run build` once in the source folder) and restart your client.

## First-run check

In any chat, type:

> Run `excel_check_environment` and show me the result.

Output tells you:
- whether Excel is installed and running,
- whether VBA Trust is enabled (Windows),
- whether macOS Automation permission is granted,
- which folders are sandboxed,
- which categories of tools work right now on your machine.

If something's off, the response includes specific fixes (links to Trust Center settings, System Preferences → Privacy → Automation, etc.).

## Common first-time issues

| Symptom | Fix |
|---|---|
| "Tools icon doesn't show 62 tools" | Server didn't load. Check Claude logs. Most often: invalid JSON in your config file. |
| "PATH_OUTSIDE_ALLOWED" error | The `.xlsx` you're working with is outside your sandboxed folders. Add the folder to `Allowed Directories` in the extension settings (or `EXCEL_ALLOWED_DIRS` in your config). |
| Tool hangs forever (Mac) | First call needs Automation permission. Switch to Excel app, then back to Claude — the permission prompt should appear. Approve once. |
| "VBA access denied" (Windows) | Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → enable **Trust access to the VBA project object model**. |
| Tool says "Requires Windows" but you're on Mac | See [PLATFORM_PARITY.md](PLATFORM_PARITY.md) for what works where. |

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more.
