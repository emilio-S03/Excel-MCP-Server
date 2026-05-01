# Platform Parity Matrix — v3.2.0 (96 tools)

Excel MCP Server runs on **Windows**, **macOS**, and **Linux**. Most tools work everywhere. Some tools depend on talking to a *running* Excel application via OS-level automation, and the support varies.

## Quick summary

| Category | Tools | Windows | macOS | Linux |
|---|---|:-:|:-:|:-:|
| File-mode (read, write, format, chart, CSV, find/replace, image, validate) | ~38 | ✅ | ✅ | ✅ |
| Live cell read/write while Excel is open | 7 | ✅ | ✅ | ❌ |
| Live cell formatting | 4 | ✅ | ✅ | ❌ |
| Live sheet management | 4 | ✅ | ✅ | ❌ |
| Live row/column ops | 4 | ✅ | ✅ | ❌ |
| Live display options (gridlines, headers, zoom, freeze, tab color) | 1 | ✅ | ✅ *unverified* | ❌ |
| Live sheet protection | 1 | ✅ | ✅ *unverified* | ❌ |
| Live calculation mode + recalc | 3 | ✅ | ✅ *unverified* | ❌ |
| Live screenshot | 1 | ✅ (range) | ✅ *unverified* (whole window only) | ❌ |
| Capability probe | 1 | ✅ | ✅ | ✅ |
| Live charts (create + style) | 2 | ✅ | ❌ planned v3.1 | ❌ |
| Live shapes (dashboard cards) | 1 | ✅ | ❌ planned v3.1 | ❌ |
| Live pivot tables | 1 | ✅ | ❌ planned v3.1 | ❌ |
| Live tables (ListObject) | 1 | ✅ | ❌ planned v3.1 | ❌ |
| Live data validation | 1 | ✅ | ❌ planned v3.1 | ❌ |
| Live conditional formatting | 1 | ✅ | ❌ planned v3.1 | ❌ |
| Live batch format | 1 | ✅ | ❌ planned v3.1 | ❌ |
| Live comments (read/add) | 2 | ✅ | ❌ planned v3.1 | ❌ |
| Live named ranges (list/create/delete) | 3 | ✅ | ❌ planned v3.1 | ❌ |
| Power Query (list/run) | 2 | ✅ | ❌ no Mac path | ❌ |
| VBA macro execution + module read/write | 5 | ✅ (Trust required) | ❌ Microsoft removed AppleScript VBA | ❌ |
| VBA Trust registry check/set | 2 | ✅ | ❌ Windows-only registry | ❌ |
| Diagnose COM connection | 1 | ✅ | ❌ use `excel_check_environment` | ❌ |

**\*unverified** = AppleScript implementation written from documented patterns but not yet validated against a real Mac with Excel 2024 — please report any issues.

## What "❌" means

Calling a Windows-only tool on macOS/Linux returns a structured `PLATFORM_UNSUPPORTED` error that tells you:

1. **Why** the tool can't run on your platform.
2. **What to use instead** (file-mode equivalent, Office Scripts link, etc.).
3. To run `excel_check_environment` for the full picture.

This is intentional: silent fallbacks lead to "it worked on Windows but did nothing on my Mac" bug reports.

## VBA on Mac

Microsoft removed the AppleScript path for invoking VBA macros from Excel for Mac years ago. There is no workaround at the OS-automation layer. Two options:

1. **Use Office Scripts** (https://learn.microsoft.com/office/dev/scripts/) — TypeScript-style macros that run in Excel for the Web and via Power Automate. Cross-platform.
2. **Run the VBA tools on a Windows machine.**

## Power Query on Mac

Excel for Mac supports Power Query in the UI but does not expose Power Query refresh/creation through AppleScript. No automation path exists. Office Scripts is again the cross-platform alternative for similar workloads.

## What's planned for v3.1

The unverified-on-Mac tools and the "planned v3.1" entries above need real Mac validation (Spike A in the upgrade plan). When that happens, the matrix updates and any patterns that don't translate to AppleScript get documented as Windows-only with their Office Scripts alternative.

## v3.2 additions

| Category | Tool | Windows | macOS | Linux |
|---|---|:-:|:-:|:-:|
| Sparklines (file-mode, OOXML manipulation) | `excel_add_sparkline`, `excel_remove_sparklines` | ✅ | ✅ | ✅ |
| Modern chart types (live mode COM only) | `excel_create_modern_chart` (waterfall/funnel/treemap/sunburst/histogram/boxWhisker), `excel_create_combo_chart` | ✅ | ❌ Office Scripts | ❌ |
| Live inspection (read existing structures) | `excel_list_charts`, `excel_get_chart`, `excel_list_pivot_tables`, `excel_list_shapes` | ✅ | ⚠️ AppleScript stubs (UNVERIFIED — coworker testing required) | ❌ |
| Formula audit (file-mode) | `excel_find_formula_errors`, `excel_find_circular_references`, `excel_workbook_stats`, `excel_list_formulas`, `excel_trace_precedents` | ✅ | ✅ | ✅ |

> **Mac coworkers:** see [MAC_VERIFICATION_NEEDED.md](MAC_VERIFICATION_NEEDED.md) for the verification checklist. The 4 v3.2 live-inspection tools (`list_charts`, `get_chart`, `list_pivot_tables`, `list_shapes`) currently throw an `UNVERIFIED_MAC_INSPECTION` error on Mac — the AppleScript implementations are stubbed and need real-Mac validation by a coworker before being unblocked.
