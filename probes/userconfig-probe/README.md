# Spike B — user_config delivery channel probe

**Purpose:** Determine which channel (`params._meta`, `params.capabilities.experimental`, `manifest.json` `user_config`, or none) Claude Desktop and Claude Code actually use to deliver runtime configuration to an MCP server.

**Why it matters:** The current excel-mcp-server expects config via a `notifications/configure` notification that neither client appears to send. Phase 1 of the upgrade implements a real config channel — but only if we know which one works.

## How to run

### Option 1: Direct node command (works in both Claude Desktop and Claude Code)

Add this to `%APPDATA%\Claude\claude_desktop_config.json` and `~/.claude.json` `mcpServers` block:

```json
"userconfig-probe": {
  "command": "node",
  "args": ["C:\\Users\\Emilio\\mcp-servers\\excel-mcp-server\\probes\\userconfig-probe\\probe.mjs"],
  "env": {}
}
```

Restart Claude Desktop and Claude Code. Then in either client:

1. Confirm the probe loaded — `tools/list` should include `probe_dump`.
2. Call `probe_dump` to print the log file location.
3. Open `~/mcp-userconfig-probe.log` (on Windows: `C:\Users\Emilio\mcp-userconfig-probe.log`).

### Option 2: .mcpb bundle (tests the user_config UI in Claude Desktop)

```bash
cd probes/userconfig-probe
npx @anthropic-ai/mcpb pack .
```

Double-click the resulting `userconfig-probe-0.0.1.mcpb` to install in Claude Desktop. The Settings → Extensions UI should expose the four `user_config` fields (string, number, bool, directories). Set non-default values, restart Claude Desktop, then read the log.

## What to look for in the log

Each entry is timestamped + labeled. Key labels:

- `INITIALIZE_PARAMS_META` — if non-null and contains `userConfig`, the spec-style `_meta` channel works.
- `INITIALIZE_PARAMS_CAPS_EXP` — if contains `userConfig`, the experimental capabilities channel works.
- `NOTIFICATION:notifications/configure` — if present, the legacy notification channel is alive.
- `NOTIFICATION:*` — any other notification method observed.
- `INITIALIZE_CLIENTINFO` — distinguishes Claude Desktop vs Claude Code in mixed logs.

## Decision matrix

| Log shows | Phase 1 implementation |
|---|---|
| `userConfig` in `_meta` or experimental | Hydrate from `InitializeRequestSchema` handler params |
| `notifications/configure` fires | Keep existing notification handler, but verify it's not just an echo |
| Nothing — only stock initialize fields | Fall back to `manifest.json` `user_config` block; document JSON-edit path for Claude Code |

## After running

Capture the relevant log excerpts in `RESULTS.md` next to this README so the decision is auditable.
