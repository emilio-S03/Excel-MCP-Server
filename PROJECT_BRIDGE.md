# Excel MCP Server - Shared Project Bridge

> **This file is the shared memory between Claude Code (terminal) and Claude Desktop.**
> Both interfaces MUST read this file at the start of every session and update it before ending.
> Last updated: 2026-03-20

---

## Current State

- **Version:** 2.0.0
- **Status:** Production-ready, 56 tools (55 base + style_chart)
- **Platform:** Windows 11 (user's machine) — AppleScript features are Mac-only, not active
- **Location:** `C:\Users\emjsa\mcp-servers\excel-mcp-server\`
- **Config:** `C:\Users\emjsa\AppData\Local\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json`
- **Build:** TypeScript → `dist/` via `npm run build`
- **Entry point:** `dist/index.js` (run via `node`)

## Tool Categories (56 total)

| Category | Count | Tools |
|----------|-------|-------|
| Read | 5 | read_workbook, read_sheet, read_range, get_cell, get_formula |
| Write | 5 | write_workbook, update_cell, write_range, add_row, set_formula |
| Format | 4 | format_cell, set_column_width, set_row_height, merge_cells |
| Sheets | 4 | create_sheet, delete_sheet, rename_sheet, duplicate_sheet |
| Operations | 3 | delete_rows, delete_columns, copy_range |
| Analysis | 2 | search_value, filter_rows |
| Charts | 2 | create_chart, **style_chart** (NEW) |
| Pivot Tables | 1 | create_pivot_table |
| Tables | 1 | create_table |
| Validation | 4 | validate_formula_syntax, validate_range, get_data_validation_info, set_data_validation |
| Advanced | 4 | insert_rows, insert_columns, unmerge_cells, get_merged_cells |
| Conditional | 1 | apply_conditional_format |
| Comments | 2 | get_comments, add_comment |
| Named Ranges | 3 | list_named_ranges, create_named_range, delete_named_range |
| Protection | 1 | set_sheet_protection |
| Calculation | 3 | trigger_recalculation, get_calculation_mode, set_calculation_mode |
| Screenshot | 1 | capture_screenshot |
| VBA | 5 | run_vba_macro, get_vba_code, set_vba_code, check_vba_trust, enable_vba_trust |
| Diagnosis | 1 | diagnose_connection |
| Power Query | 2 | list_power_queries, run_power_query |
| Batch Format | 1 | batch_format |
| Display | 1 | set_display_options |
| Shapes | 1 | add_shape |

## Architecture Quick Reference

```
src/
├── index.ts              ← Entry point, tool registration, routing
├── constants.ts          ← Tool annotations, error messages, defaults
├── types.ts              ← TypeScript interfaces
├── schemas/
│   └── index.ts          ← Zod validation for all 56 tools
└── tools/
    ├── helpers.ts         ← Shared utils (load/save, path validation)
    ├── read.ts            ← Read operations
    ├── write.ts           ← Write operations (with COM fallback warnings)
    ├── format.ts          ← Cell formatting
    ├── sheets.ts          ← Sheet management
    ├── operations.ts      ← Data operations
    ├── analysis.ts        ← Search/filter
    ├── charts.ts          ← Chart creation
    ├── pivots.ts          ← Pivot tables
    ├── tables.ts          ← Table formatting
    ├── validation.ts      ← Validation tools
    ├── advanced.ts        ← Insert rows/cols, merge
    ├── conditional.ts     ← Conditional formatting
    ├── vba.ts             ← VBA macro execution with safety scanning + error parsing
    ├── diagnose.ts        ← excel_diagnose_connection tool
    ├── excel-live.ts      ← Platform dispatcher (Windows COM / Mac AppleScript)
    ├── excel-powershell.ts← Windows COM automation (PowerShell), retry logic, VBA wrapper
    └── excel-applescript.ts← Mac-only collaborative mode
```

## Known Issues & Limitations

- AppleScript collaborative mode is Mac-only — not functional on Windows
- ExcelJS has limited native chart support (creates metadata placeholders)
- Pivot tables are manually calculated (no native ExcelJS pivot support)
- Macros and custom XML may not survive round-trip editing
- No `allowedDirectories` currently set in config (unrestricted file access)
- COM connection drops when VBA macros have unhandled runtime errors showing modal dialogs — **mitigated** by Fix 1 (VBA error wrapper) and Fix 4 (error categorization) as of 2026-03-13
- Write tools silently fell back to file-based editing when COM failed — **mitigated** by Fix 3 (fallback warnings) as of 2026-03-13
- VBA COM tools do NOT survive a crashed macro in the same session — restarting Excel is the only recovery. Fix 1 prevents most crashes, but if one slips through, session is dead for VBA tools. **Hard constraint, not fixable server-side.**
- `Chr()` with values >255 in VBA strings crashes COM connection — **mitigated** by VBA scanner warning as of 2026-03-13
- `On Error Resume Next` in VBA macros masks failures, making them invisible to the MCP error wrapper — **mitigated** by VBA scanner warning as of 2026-03-13
- ~~`excel_create_chart` is basic — no native support for series colors, axis formatting, data labels, or background fills.~~ **RESOLVED** by `excel_style_chart` (tool #56) — 2026-03-13
- ~~`excel_add_shape` has no text auto-sizing — text overflows if font is too large for shape dimensions~~ **RESOLVED** by `autoSize` parameter — 2026-03-13
- ~~VBA `&` string concatenation arrives as `&amp;` via MCP protocol, breaking all VBA macros that use string concat~~ **RESOLVED** by HTML entity decoding in `setVbaCodeViaPowerShell` — 2026-03-16
- ~~`excel_style_chart` data labels are all-or-nothing per series — no way to hide individual labels on small/zero segments in stacked charts, causing overlapping labels~~ **RESOLVED** by `hideBelow` parameter on `dataLabels` — 2026-03-16

## Backlog / Future Work

- [x] Named ranges support (implemented)
- [ ] Image insertion
- [x] Comments/notes management (implemented)
- [x] Protection/security features (sheet protection implemented)
- [ ] Performance optimization for large files
- [ ] Additional chart customization
- [x] Windows collaborative mode (COM/PowerShell — fully operational with 55 tools)
- [ ] Add `allowedDirectories` to config for security
- [x] Test COM resilience fixes end-to-end (Fix 1-4 — validated by Desktop 2026-03-13)
- [x] Enhanced chart tool — `excel_style_chart` (tool #56) — series colors, data labels, axis formatting, chart/plot area fills, legend, title, size. One call replaces VBA write+run cycle. Implemented 2026-03-13.
- [x] Shape text auto-sizing — `autoSize` parameter on `excel_add_shape` text object: `'shrinkToFit'` (auto-shrinks text), `'shapeToFitText'` (grows shape). Implemented 2026-03-13.

---

## Session Log

> Each session should add an entry below. Format:
> `### [Date] — [Interface] — [Summary]`
> Keep entries brief. Move old entries to ARCHIVE section when this file exceeds 150 lines.

### 2026-03-16 — Claude Code — Overlapping label fix + VBA encoding fix
- **FIX A — Per-point label visibility (`hideBelow`):** New `hideBelow` parameter on `dataLabels` in `excel_style_chart`. When set, iterates through each data point in the series and hides the label if `abs(value) < threshold`. Example: `hideBelow: 0.05` hides labels on segments smaller than 5% in a percentage stacked bar chart. Prevents the overlapping "0%" issue when cycling scenarios. Schema (`schemas/index.ts`), TypeScript interface, and PowerShell code (`excel-powershell.ts`) all updated.
- **FIX B — HTML entity decoding in VBA code:** Added `decodeHtmlEntities()` function that runs before `escapePowerShellString()` in `setVbaCodeViaPowerShell`. Decodes `&amp;` → `&`, `&lt;` → `<`, `&gt;` → `>`, `&quot;` → `"`, `&#39;`/`&#x27;` → `'`. This unblocks VBA macros that use `&` for string concatenation — Desktop can now write `"text" & "more"` without it becoming `"text" &amp; "more"` in the CodeModule.
- **Root cause of Desktop's label formatting destruction:** Desktop couldn't write a VBA fix (encoding bug), so it used raw PowerShell to manipulate labels — which overwrote NumberFormat from "0%" to "$#,##0" and broke the font. Both server fixes prevent this class of failure.
- **Desktop testing needed:**
  1. `excel_style_chart` with `dataLabels: { show: true, numberFormat: "0%", hideBelow: 0.01 }` on each series of the stacked bar — cycle all 4 scenarios, verify 0%/tiny segments have no labels
  2. `excel_set_vba_code` with VBA code containing `&` string concatenation — verify it compiles and runs
- Build passes cleanly.

### 2026-03-13 — Claude Code — Bug fixes from Desktop testing (3 bugs)
- **BUG 1 FIX — DataLabels crash in `excel_style_chart`:** Root cause was `$ErrorActionPreference = 'Stop'` making non-critical DataLabels property assignments terminate the entire call. Fix: (a) Re-fetch the series object after setting `HasDataLabels = $true` (COM object may not be initialized immediately), (b) Wrap each DataLabels property (NumberFormat, Font.Size, Font.Color, Position) in individual `try/catch` blocks so one failure doesn't kill the styling call.
- **BUG 2 FIX — Date serial numbers on x-axis:** The NewSeries-Array tier now detects date-formatted cells by checking `NumberFormat` of column 1 cells (matches patterns containing d/m/y but not numeric formats). For date cells, uses `$cell.Text` (Excel's formatted display value like "1/30") instead of `$cell.Value2` (raw serial number like 46052). Non-date cells still use `[string]$cv`.
- **BUG 3 FIX — `excel_set_vba_code` overwrites module:** New `appendMode` parameter (boolean, default false). When true, appends the new code after existing module content with a blank line separator instead of clearing the module first. Schema, dispatcher, entry point, and PowerShell all updated.
- Build passes cleanly. All 3 fixes ready for Desktop testing.

### 2026-03-13 — Claude Code — Native Chart Styling + Shape Auto-Sizing + Chart Data Binding Fix (v4)
- **New tool: `excel_style_chart` (tool #56)** — Styles existing charts via COM. One call replaces VBA write-macro + run-macro cycle.
- **Enhanced `excel_add_shape`** — New `autoSize` parameter in text object. Desktop confirmed working.
- **Fixed `excel_create_chart` data binding (4 iterations):**
  - v1: Replaced `Shapes.AddChart2` with `ChartObjects().Add()` — `SetSourceData` silently failed (no exception)
  - v2: Added `$ErrorActionPreference = 'Stop'` + 3-tier try/catch — but `SetSourceData` returns S_OK with 0 series (no exception to catch)
  - v3: Desktop confirmed: all tiers silently "succeed" but produce 0 series. Root cause was relying on exceptions when the COM methods don't throw.
  - **v4 (current — based on Desktop's diagnostic recommendations):** Complete rewrite. Now checks `SeriesCollection.Count` after each tier instead of relying on exceptions:
    1. **Tier 1:** `SetSourceData($srcRange)` → check count
    2. **Tier 2:** `NewSeries()` with range references → check count (clears partial series on failure)
    3. **Tier 3:** `NewSeries()` with extracted array values (`[double]`/`[string]` cast per cell) → check count
    4. **Tier 4:** Hardcoded `@(1,2,3,4,5)` test — if this fails, the chart object itself can't accept series (Dashboard sheet issue)
    5. If all 4 tiers produce 0 series: **deletes the empty chart** and throws `CHART_BIND_FAILED` with all tier errors + range dimensions
  - **Return value now includes diagnostics:** `bindingTier`, `seriesCount`, `dataRows`, `dataCols` — so Desktop knows which tier worked (or all failed)
  - **`dataSheetName` parameter** — place chart on Dashboard, pull data from another sheet
- **Desktop testing needed:**
  1. `excel_create_chart` with Dashboard data — will now either bind data or throw `CHART_BIND_FAILED` with tier-by-tier errors (never silently return success with 0 series)
  2. `excel_create_chart` with `dataSheetName: "Net_Worth"` — isolates Dashboard sheet issues
  3. Check `bindingTier` in response to see which approach worked
  4. If Tier 4 (hardcoded) also fails → confirms Dashboard sheet itself rejects chart series (possible protection or corruption)

### 2026-03-13 — Claude Code — COM Resilience Fixes (4 fixes)
- **Fix 4:** Error categorization in `execPowerShellWithRetry` — timeout, COM unreachable, VBA trust, HRESULT errors now have actionable messages pointing to `excel_diagnose_connection`
- **Fix 2:** New `excel_diagnose_connection` tool (tool #55) — 6-step diagnostic: process → COM → responsive → file open → VBA trust → registry. New files: `diagnose.ts`, changes to `excel-powershell.ts`, `excel-live.ts`, `schemas/index.ts`, `index.ts`
- **Fix 3:** Write path fallback warning — all 4 write functions (`updateCell`, `writeRange`, `addRow`, `setFormula`) now warn when falling back to ExcelJS while Excel is running
- **Fix 1:** VBA error handler auto-injection — `runVbaMacroViaPowerShell` creates a temp VBA module with `On Error GoTo ErrHandler` wrapper to catch unhandled errors before they show modal dialogs. Falls back to direct Run if no VBProject access. `vba.ts` parses structured VBA error responses
- Build passes cleanly. **Needs end-to-end testing** (see test scenarios in plan)
- Updated bridge file: corrected tool count to 55, updated architecture map, known issues, backlog

### 2026-03-13 — Claude Code — Bridge file created
- Created `PROJECT_BRIDGE.md` as shared memory between Claude Code and Claude Desktop
- Updated `CLAUDE.md` with bridge file protocol
- Created Desktop skill for bridge file awareness
- Consolidated current state from PROGRESS.md and FEATURE_SUMMARY.md

---

## Archive

_Move old session log entries here when the file gets long._

### 2026-03-13 — Claude Desktop — Phase 2 build + COM resilience testing
- **Tested Code's Fix 1-4:** After Excel restart, all VBA tools (get/set/run) recovered. `excel_diagnose_connection` worked perfectly — 6/6 checks passed. Confirmed the Chr(9888) crash root cause: replaced with ASCII `[!]`, macro ran clean.
- **Built Debt_Tracker sheet:** Current balances pulled from Balances_Norm, baseline comparison (Jan 30 vs Mar 12), avalanche payoff projections with formulas. Warning for Chase Freedom 0% expiry.
- **Built Net_Worth sheet:** 42 daily snapshots (Jan 30 – Mar 12), assets vs liabilities breakdown, daily change column, summary section. Starting NW: -$5,076, Current: -$8,591.
- **Built Spending_Analysis sheet (earlier this session):** 25 months x 44 categories from Transactions_Clean, monthly totals + income row.
- **Known issue confirmed:** VBA COM tools do NOT survive a crashed macro in the same session. Restarting Excel is the only recovery. Code's Fix 1 (error wrapper) should prevent the crash in the first place, but if it fails, the session is dead for VBA tools.
- **Anti-pattern documented:** Never use Chr() with values >255 in VBA strings sent via COM. Use ASCII alternatives.

### 2026-03-13 — Claude Desktop — Dashboard rebuild + chart data binding failure (CRITICAL)

**Status:** Shapes work perfectly. Charts still broken. COM dropped at end of session.

**What was built successfully:**
- Title bar, 4 KPI cards (with live values: -$8,591, -$15,369, $9,104, Oct 2026), 3 chart card containers — all via VBA `Shapes.AddShape`. `autoSize` on `excel_add_shape` confirmed working for shrinkToFit.
- Canvas: navy background (#1B2A4A), gridlines hidden, headers hidden, column widths set.
- All shapes survive nuke/rebuild cycles reliably.

**What is still broken — `excel_create_chart` produces 0 series:**
- `excel_create_chart` returns `success: true` with message "Real line chart created" — but the chart has **0 series**. Every time.
- Tested with `dataRange: "AH1:AI15"` on Dashboard (dates + numbers), and without `dataSheetName`.
- `excel_style_chart` then crashes at `$chart.SeriesCollection(1)` → "Parameter not valid" because there's nothing to style.
- **VBA manual testing confirms the same:** `SetSourceData`, `NewSeries` with range refs, and `NewSeries` with direct assignment ALL fail with "Parameter not valid" when targeting charts on the Dashboard sheet.
- The v3 three-tier fallback does NOT appear to be firing — the tool returns success with no fallback indicators, yet series count is 0. Hypothesis: `$ErrorActionPreference = 'Stop'` may not be reaching the chart creation PowerShell scope, OR the chart creation code path returns success before the fallback code runs.

**Diagnostic evidence:**
- VBA `ChkChart` macro confirmed: `Chart.SeriesCollection.Count = 0` on every chart created by `excel_create_chart`.
- VBA `ChartDiag3` attempted manual `SetSourceData` and `NewSeries` on the same chart objects — both fail with "Parameter not valid".
- Data in AH1:AI15 is confirmed valid (dates in AH, doubles in AI, headers in row 1). Data in AH20:AI24 (debt) and AH30:AI38 (spending) also confirmed valid.
- Dashboard cell writes via VBA to BA/AZ columns sometimes don't persist (race condition? stale COM handle?). Writes to _Config sheet untested.
- COM connection eventually dropped after ~25 tool calls in the session.

**Root cause hypothesis:**
The `excel_create_chart` PowerShell is creating charts via `ChartObjects().Add()` which creates an **empty** chart object. The subsequent `SetSourceData` or `NewSeries` calls are silently failing or not executing at all. The function returns success based on the ChartObject creation, not on data binding. The three-tier fallback may be in a code path that isn't reached, or `$ErrorActionPreference = 'Stop'` isn't propagating into the chart data binding scope.

**Recommended fix for Claude Code:**
1. **Add explicit series count validation** after data binding — if `SeriesCollection.Count -eq 0` after all 3 tiers, throw an error instead of returning success.
2. **Add diagnostic output** to the chart creation return: include `seriesCount`, `fallbackTier` (which tier succeeded), and if all fail, the specific error from each tier.
3. **Test the array extraction tier (tier 3) in isolation** — write a standalone PowerShell snippet that creates a chart via `ChartObjects().Add()`, then does `NewSeries()` with hardcoded `@(-5076, -5140, -9953)` array values. If THAT works, the issue is in tier 1-2 range marshaling. If it fails, the issue is deeper (maybe chart type incompatibility or a Dashboard sheet-level protection).
4. **Consider testing on a non-Dashboard sheet** — create a chart on `_Config` or a new scratch sheet with the same data range to rule out Dashboard-specific issues.
5. The `dataSheetName` parameter was not successfully tested — the first test used `dataSheetName: "Dashboard"` (same sheet) which doesn't help isolate the issue. Need to test with `dataSheetName: "Net_Worth"` pulling from a clean sheet.

**Dashboard state after session:**
- All shapes were nuked at end (FullNuke ran). Dashboard is a blank navy canvas with data in hidden columns (AA:AI).
- Module1 contains the last macro written (`FullNuke`). All prior macros (BuildAllShapes, etc.) were overwritten.
- No charts exist.
- To rebuild: run BuildAllShapes VBA (needs to be re-written to Module1), then fix chart creation.


### 2026-03-13 — Claude Desktop — Dashboard COMPLETE + bug reports for Code

**Status:** Dashboard fully functional. Charts render with real data. Layout rebuilt at proper scale.

**What works now:**
- `excel_create_chart` v4 (NewSeries-Array tier) **confirmed working** — all 3 charts bind data successfully. Response includes `bindingTier: "NewSeries-Array"`, `seriesCount: 1`. The v3 fix from Code resolved the 0-series issue.
- `excel_style_chart` works for: series fill color, chart/plot area fills, axis formatting (font, color, size, numberFormat, gridline color), legend hide, title hide, chart size/position.
- VBA macros work reliably for: shape creation, chart repositioning (`ZOrder msoBringToFront`), x-axis label override (date serial → formatted "m/d"), data label application.
- Full dashboard layout: title bar, 4 KPI cards (accent stripes, labels, values), 3 chart cards with titles, all proportionally sized at 936pt usable width.

**Dashboard current state:**
- Sheet: `Dashboard` in `Finances.xlsm`
- Zoom: 100%, gridlines hidden, headers hidden
- Shapes: Title bar + 4 KPI cards + 3 chart card backgrounds + labels (all via VBA `RebuildDashboard` macro)
- Charts: 3 ChartObjects (NW line, Debt bar, Spending bar) styled and positioned inside cards via `StyleAndPositionCharts` macro
- Data source: hidden cols AH:AI on Dashboard sheet (NW=AH1:AI15, Debt=AH20:AI24, Spending=AH30:AI38)
- KPI source: AB1:AB4 (formulas linking to Net_Worth and Debt_Tracker sheets)
- Module1 contains: `StyleAndPositionCharts` (last written macro — overwrites previous)

**BUG 1 — `excel_style_chart` DataLabels crash (NEEDS FIX)**
- **Trigger:** Passing `dataLabels: { show: true, numberFormat: "$#,##0", fontColor: "#FFFFFF", fontSize: 8 }` in the `series` array
- **Error:** `PropertyAssignmentException: The property 'NumberFormat' cannot be found on this object` and `The property 'Size' cannot be found on this object`
- **Root cause:** `$ErrorActionPreference = 'Stop'` makes non-critical DataLabels property assignments terminate the entire styling call. The DataLabels COM object properties (`NumberFormat`, `Font.Size`, `Font.Color`) may use different property paths than expected, OR the DataLabels object isn't fully initialized when accessed immediately after `.HasDataLabels = $true`.
- **Workaround used:** Applied data labels via VBA macro instead of `excel_style_chart`. Works perfectly in VBA.
- **Recommended fix:** Wrap each DataLabels property assignment in its own try/catch block inside `excel-powershell.ts` so one failure doesn't kill the whole call. Also add a small delay or re-fetch after setting `HasDataLabels = $true` before accessing properties. Test: `$series.DataLabels.NumberFormat` vs `$series.DataLabels.Item(1).NumberFormat` — the collection accessor may differ.

**BUG 2 — Net Worth chart x-axis shows serial date numbers**
- **Trigger:** `excel_create_chart` with date values in the XValues column (column AH contains Excel date serials like 46052, 46055...)
- **Behavior:** The NewSeries-Array tier extracts cell values as raw doubles, so dates become serial numbers (46052 instead of "1/30"). Chart displays "46052" on x-axis.
- **Workaround used:** VBA macro overrides XValues with a string array of `Format(date, "m/d")` labels after chart creation.
- **Recommended fix:** In the array extraction tier, detect date-formatted cells (check `$cell.NumberFormat` for date patterns like "m/d/yyyy", "mm/dd/yy", etc.) and convert to formatted string instead of raw double. Or add a `dateFormat` parameter to `excel_create_chart` that applies to XValues.

**BUG 3 (minor) — `set_vba_code` overwrites entire module**
- Not a bug per se, but a constraint that bit us multiple times. Writing a new macro to Module1 deletes all previous macros. Desktop workaround: always include ALL needed macros in a single `set_vba_code` call.
- **Suggestion:** Consider adding an `appendMode` parameter that appends the Sub/Function to existing module code instead of replacing it.

**Sheet3 deletion:** Already gone (deleted in a prior session or never existed in current file state).

**Next planned work:**
- Dashboard polish (tighten layout, clean remaining visual issues)
- Format data sheets (Spending_Analysis, Debt_Tracker, Net_Worth) with design system


### 2026-03-20 — Claude Code — Capture index push + bridge sync
- Pushed `capture-index/index.json` (15 entries) to GitHub repo `AOLUX003/claude_config` — first-ever population of the learning audit dedup index.
- Verified EXCEL_MCP_LEARNINGS.md sections 9-11 (7 new entries) — clean, no append artifacts.
- Caught and reverted encoding corruption in `learning-audit/SKILL.md` (Desktop's read/write mangled `→` to `â†'`). Only clean capture-index committed.
- Read AUDIT_IMPLEMENTATION_SYNC.md — aware of excel-design-system v2.0 hard-stop gate, prompt-guard v3.0 Pattern 6, memory edits #25/#26.
- Updated bridge file: header date, tool count consistency.
- **No MCP server changes needed.** Audit confirmed all failures were behavioral, not tooling.

### 2026-03-20 — Claude Desktop — Thread audit implementation (no server changes)
- Completed implementation of all Desktop-scope items from GS-ForecastCalc thread audit (1,927 lines, 40+ failures).
- **No MCP server bugs or feature requests from this audit.** All findings were behavioral (Claude skipping existing protocols), not tooling gaps.
- Updated skills: excel-design-system v2.0 (hard-stop enforcement gate), prompt-guard v3.0 (Pattern 6).
- Updated EXCEL_MCP_LEARNINGS.md in Obsidian vault: added sections 9-11 (7 new entries — chart ops, formula direction, column addressing).
- Full sync doc at: `C:\Users\emjsa\Desktop\EMJCLaude\projects\gs-forecastcalc\AUDIT_IMPLEMENTATION_SYNC.md`
- **For Code:** Item T remains — push `claude_config/capture-index/index.json` to GitHub. Handoff at `C:\Users\emjsa\Desktop\CLAUDE_CODE_AUDIT_HANDOFF.md`.
