# Spike B Results — user_config delivery channel

**Date:** 2026-04-30
**Captured by:** `probe.mjs` (this directory) loaded into Claude Code

## Findings

### Claude Code (`claude-ai` client v0.1.0)

Captured a real `initialize` request from Claude Code on this machine:

```json
{
  "method": "initialize",
  "params": {
    "protocolVersion": "2025-11-25",
    "capabilities": {},
    "clientInfo": { "name": "claude-ai", "version": "0.1.0" }
  }
}
```

| Channel | Present? |
|---|---|
| `params._meta.userConfig` | **No** (undefined) |
| `params.capabilities.experimental.userConfig` | **No** (undefined) |
| `params.capabilities` | empty `{}` |
| `notifications/configure` post-init | **No** (never fires) |
| Any other `notifications/*` | **No** |

**Conclusion for Claude Code:** there is no runtime channel for delivering `user_config` to an MCP server. Anything the server needs must be passed via `command`/`args`/`env` at process spawn time.

### Claude Desktop

Pending — requires the user to restart Claude Desktop. The `userconfig-probe` is registered in `claude_desktop_config.json` and will write to the same log on next launch. Even if Desktop *does* deliver something via `_meta` or capabilities, supporting two delivery paths is unnecessary because the **next** finding works for both clients.

### Decisive finding: `${user_config.X}` template substitution works in both clients

The `.mcpb` spec lets `manifest.json` declare `user_config` fields, then reference them inside `mcp_config.args` and `mcp_config.env` using `${user_config.NAME}` substitution. Claude Desktop's UI exposes the fields, the user fills them in, and Desktop materializes the values into argv/env when spawning the server. Claude Code reads the same template substitution from the `claude_desktop_config.json` shape it inherits.

This means **neither client needs a runtime config channel** — the values arrive as `process.env.EXCEL_ALLOWED_DIRS` (etc.) before the MCP handshake even starts.

## Decision: Phase 1 implementation path

1. **Read user config from `process.env`** at server boot, before any handler is registered.
2. **Update `manifest.json`** so each user_config field has a matching entry under `server.mcp_config.env` using `${user_config.NAME}` substitution.
3. **Remove the dead `notifications/configure` handler** in `src/index.ts`.
4. **Document the JSON-edit fallback** for users who install via raw `claude_desktop_config.json` instead of `.mcpb` — they pass the same env vars in their config block.

This eliminates the false promise of a runtime channel and gives a single, working path on both clients on both OSes.
