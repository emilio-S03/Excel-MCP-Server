# Excel MCP Server — Bridge Protocol for Claude Desktop

## The Problem

The user maintains a custom Excel MCP Server — a TypeScript project that gives Claude Desktop the ability to read, write, format, and analyze Excel spreadsheets. It has 34 tools and lives at `[project-root]\`.

This project is actively developed and maintained using **two separate Claude interfaces:**

- **Claude Code (terminal)** — used for writing code, debugging, building, running tests, and making changes to the server's TypeScript source files.
- **Claude Desktop (this interface)** — used for testing the Excel tools in real conversations, discovering bugs during use, requesting new features, and working with Excel files directly.

**The pain point:** These two interfaces have no shared memory. Every time the user switches between them, context is lost. Claude Code doesn't know what bugs Desktop discovered during testing. Desktop doesn't know what Code just fixed or changed. The user ends up re-explaining the same things, work gets duplicated, bugs get re-introduced, and neither side has a clear picture of the project's current state.

**The gap this solves:** There is no built-in way for Claude Code and Claude Desktop to share context. They are completely separate sessions with separate memory. The user — who is not a developer and works by giving plain-English instructions — shouldn't have to be the "translator" between two AI interfaces working on the same project.

## The Solution

A single shared file acts as a bridge:

**`[project-root]\PROJECT_BRIDGE.md`**

This file contains:
- **Current State** — version, status, platform, build info
- **Tool inventory** — all 34 tools organized by category
- **Architecture map** — file structure and what each source file does
- **Known Issues** — bugs, limitations, things that don't work
- **Backlog** — planned future work, feature requests
- **Session Log** — a running record of what each interface did and when

Both Claude Code and Claude Desktop read this file at the start of every session and update it before ending. This way, whichever interface the user opens next always knows what happened last.

## What You (Claude Desktop) Must Do

### At the START of every session involving the Excel MCP Server:

1. **Read the bridge file** — Use your file reading capability to read `[project-root]\PROJECT_BRIDGE.md`
2. **Check the session log** — See what Claude Code (or a previous Desktop session) did most recently
3. **Note any known issues** — So you don't run into the same bugs or re-report something already tracked

### DURING your session:

4. **If you discover a bug** while using the Excel tools — note it mentally, you'll record it at the end
5. **If the user requests a feature or change** — note it for the backlog
6. **If something that was broken is now working** — note it so the issue can be marked resolved

### At the END of your session (or when the user is wrapping up):

7. **Update the bridge file** with any changes:
   - Add or remove items from "Known Issues" based on what you found
   - Add items to "Backlog" if the user requested new features
   - Mark backlog items as done (`[x]`) if they were completed
   - Update "Current State" if version, tool count, or status changed
   - **Add a session log entry** in this format:
     ```
     ### [Date] — Claude Desktop — [Brief summary of what happened]
     - Bullet points of key actions, discoveries, or decisions
     ```

## Important Context About the User

- **Non-developer.** They use plain-English prompts to direct Claude. Don't use jargon or assume coding knowledge.
- **"Vibe coder."** They describe what they want, and Claude builds it. They evaluate results by whether things work, not by reading code.
- **Windows 11.** The Excel MCP server runs on Windows. Any AppleScript/Mac-related features in the codebase are inactive and irrelevant.
- **Claude Desktop is the Windows Store version.** Config is at the Windows Store app data path (not the usual AppData\Roaming path). Check your system for the exact location.

## Important Context About the Project

- **55 Excel tools** — reading, writing, formatting, charts, pivot tables, conditional formatting, validation, and more
- **TypeScript + ExcelJS** — the server is built in TypeScript using the ExcelJS library for Excel file manipulation
- **MCP protocol** — the server communicates with Claude Desktop via stdio transport using the MCP SDK
- **The `CLAUDE.md` file** in the project root contains full technical details (architecture, build commands, how to add tools). If you need to guide Claude Code on a fix, reference that file.
- **Design system files** exist in `skills/` — `EXCEL_ADVANCED_DESIGN_REFERENCE.md` and `excel-design-system.md` for dashboard styling

## When the User Reports a Problem

If the user says something like "this tool isn't working" or "I got an error":

1. Check `PROJECT_BRIDGE.md` known issues — is this already tracked?
2. If it's new, try to reproduce it and gather details (error message, what file, what tool)
3. Add it to known issues in the bridge file
4. Tell the user: "I've logged this in the bridge file so Claude Code can see it and fix it next time you open the terminal."

## When the User Wants a New Feature

1. Check the backlog in `PROJECT_BRIDGE.md` — is it already listed?
2. If not, add it to the backlog with a clear description
3. Tell the user: "I've added this to the project backlog. Claude Code will see it next time and can implement it."

## Three Files That Were Created

| File | Purpose | Who uses it |
|------|---------|-------------|
| `PROJECT_BRIDGE.md` | The shared memory file — current state, issues, backlog, session log | Both Claude Code and Claude Desktop |
| `CLAUDE.md` (updated) | Technical instructions for Claude Code, now includes bridge protocol | Claude Code only |
| `skills/excel-mcp-bridge-protocol.md` (this file) | Full context and instructions for Claude Desktop | Claude Desktop only |
