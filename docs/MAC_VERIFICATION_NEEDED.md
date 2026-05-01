# Mac Verification Needed

## Why this exists

The Excel MCP Server's macOS code paths were authored on a Windows-only dev machine using the documented Excel-for-Mac AppleScript dictionary, but they have **not** been validated against a real Mac running Excel. Before any team relies on these tools in production, a Mac coworker needs to run them against real Excel for Mac and report Pass/Fail. Treat the AppleScript-backed tools as "best effort, unverified" until that happens.

## What you need to test on (prerequisites)

- macOS 13+ (Ventura or newer)
- Microsoft Excel for Mac 2021 or 2024 (perpetual or M365)
- Claude Desktop installed
- Excel MCP Server v3.x installed (double-click the `.mcpb` bundle)
- macOS Automation permission granted to Claude Desktop for Microsoft Excel
  (System Settings -> Privacy & Security -> Automation -> Claude -> enable Microsoft Excel)

## Pre-flight check

Before running any of the tests below, ask Claude:

> "Run `excel_check_environment` and show me the JSON."

Verify the report contains:

- `platform: "darwin"`
- `excel.installed: true`
- `automationPermission.granted: true`

If `automationPermission.granted` is `false`, fix that first — every other test will hang otherwise.

## Tools to verify

For each row, copy the prompt into Claude Desktop, then mark Pass/Fail. The "function" column is the source-code function backing the tool — useful when filing a failure report. All Mac functions live in `src/tools/applescript-extended.ts`.

| Tool name | What it should do | Test prompt | Expected result | Pass/Fail |
|---|---|---|---|---|
| `excel_set_display_options` (`setDisplayOptionsViaAppleScript`) | Toggle gridlines / headers, set zoom, freeze pane, set tab color | "Open `~/Documents/test.xlsx` in Excel. Then on Sheet1: hide gridlines, hide row/column headers, zoom to 90%, freeze panes at B2, and set the tab color to `#FF0000`." | Sheet1 has no gridlines, no headers, 90% zoom, A1:A and row 1 frozen, red tab | |
| `excel_set_sheet_protection` (`setSheetProtectionViaAppleScript`) | Protect/unprotect a sheet, optional password | "In `~/Documents/test.xlsx`, protect Sheet1 with password `hunter2`. Then try to type into A1 in Excel." | Excel blocks the edit; unprotect with same password restores editing | |
| `excel_get_calculation_mode` (`getCalculationModeViaAppleScript`) | Read the app-wide calculation mode | "What is Excel's current calculation mode?" | Returns `automatic`, `manual`, or `semiautomatic` | |
| `excel_set_calculation_mode` (`setCalculationModeViaAppleScript`) | Set app-wide calc mode | "Set Excel's calculation mode to `manual`, then back to `automatic`." | Excel -> Preferences -> Calculation reflects each change | |
| `excel_trigger_recalculation` (`triggerRecalculationViaAppleScript`) | Force recalc (`calculate` or `calculate full`) | "In `~/Documents/test.xlsx`, put `=NOW()` in A1, then trigger a full recalculation." | A1 timestamp updates to current time | |
| `excel_capture_screenshot` (`captureScreenshotViaAppleScript`) — Mac captures the whole window, not a range | "Open `~/Documents/test.xlsx`, then capture a screenshot of Sheet1 and save it to `~/Desktop/excel-shot.png`." | PNG file exists at the path and shows the Excel window contents | |
| `excel_export_pdf` (Mac path uses AppleScript `save as PDF`, see `src/tools/pdf-export.ts`) | Export a workbook/sheet to PDF | "Export `~/Documents/test.xlsx` to `~/Desktop/test.pdf`." | PDF exists at exact `outputPath` (verify the path — see gotcha below) | |

If you discover any other tool that reports `platform.unsupported` errors but you expect to work, add a row and test it the same way.

## What to do when a test fails

For every failure, capture the following and send it back (GitHub issue or DM Emilio):

1. The exact prompt you gave Claude
2. The exact JSON tool call Claude made (visible in Claude Desktop's "Used tool" expander)
3. The exact error string Claude reported
4. The relevant section of `docs/PLATFORM_PARITY.md` for context
5. The AppleScript that the tool tried to run — copy from the function source in `src/tools/applescript-extended.ts`
6. The output of running the failing AppleScript directly from Terminal, e.g.:

   ```
   osascript -e 'tell application "Microsoft Excel" to get calculation'
   ```

   This isolates whether the failure is in AppleScript itself or in the MCP harness wrapping it.

## Common Mac-specific gotchas

- **First call hangs forever** — Automation permission was never granted. Open System Settings -> Privacy & Security -> Automation, find Claude, enable the Microsoft Excel checkbox, then retry.
- **App name confusion** — AppleScript targets `"Microsoft Excel"`, not `"Excel for Mac"`. If you see `application "Excel" is not running`, that's the wrong identifier somewhere.
- **Sandbox / file access** — macOS may block Excel from writing to certain folders (Desktop, Downloads). Test inside `~/Documents` first; if that works but Desktop fails, it's a sandbox grant issue, not the MCP server.
- **VBA tools are intentionally Windows-only** — `excel_run_vba_macro`, `excel_get_vba_code`, `excel_set_vba_code` will throw a friendly error on Mac pointing you to Office Scripts. **That is correct behavior, not a failure.** Do not log it as a bug.
- **PDF export path** — the AppleScript `save as PDF` command can be quirky about where the file lands if not given an absolute POSIX path. After running the tool, verify the PDF actually exists at the `outputPath` you specified (not in the workbook's source folder).
- **Screenshot scope** — the Mac implementation captures the **entire Excel window**, not a specific cell range. This is a known limitation versus the Windows version. Don't file it as a bug — file it as "Mac-only: range-scoped screenshot would require a separate AppleScript path."

## Self-heal prompt template

When something fails, paste this into Claude Desktop (filling in the angle-bracketed parts):

> Tool `<tool name>` failed on my Mac with error `<exact error text>`.
> Read `src/tools/applescript-extended.ts`, find the `<functionName>` function,
> and propose an AppleScript fix. Output:
> (1) the diff,
> (2) the exact `osascript` command I should run from Terminal to verify the fix in isolation,
> (3) a `node:test` case under `test/` that Emilio can run on Windows to confirm the tool registration didn't break.

Claude has the source available and can usually triangulate the dictionary mismatch from the error text alone.

## Reporting back

When you've worked through the table:

1. Copy the Pass/Fail column (with tool names) into a Slack DM to Emilio.
2. Attach any error transcripts and the Terminal `osascript` output you gathered for each failure.
3. Note the macOS version and Excel build (`Excel -> About Microsoft Excel`) — version skew is the #1 cause of dictionary differences.
