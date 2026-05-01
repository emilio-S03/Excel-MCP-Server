/**
 * Extended AppleScript implementations for Mac parity (v3.0.0 Batch A).
 *
 * IMPORTANT: These were authored from the Excel-for-Mac AppleScript dictionary
 * documentation but have NOT yet been validated against a real Mac with Excel
 * 2024 (Spike A is deferred until Mac access is available — see plan
 * `i-want-you-to-resilient-sonnet.md` Phase -1). Treat as best-effort and
 * verify on a real Mac before relying on for production work.
 *
 * Conservative scope — only patterns that are well-documented across multiple
 * Excel-for-Mac versions:
 *   - displayOptions: standard `tell active window` + property assignments
 *   - sheetProtection: standard `protect` / `unprotect` commands
 *   - calculationMode: app-level `calculation` enum
 *   - triggerRecalculation: `calculate` command
 *
 * Other Mac-only stubs return actionable errors via platform-errors.ts so
 * users see "use file-mode tool X" or "use Office Scripts" instead of
 * "Requires Windows".
 */
import { exec } from 'child_process';
import { promisify } from 'util';
import { basename } from 'path';

const execAsync = promisify(exec);
const APPLESCRIPT_TIMEOUT = 8000;

function escapeAS(s: string): string {
  return s.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

async function runOsascript(script: string, timeoutMs = APPLESCRIPT_TIMEOUT): Promise<string> {
  // Use -e per-line so multi-line scripts pass through the shell cleanly.
  const lines = script.split('\n').filter((l) => l.trim().length > 0);
  const args = lines.map((l) => `-e "${l.replace(/"/g, '\\"')}"`).join(' ');
  const { stdout } = await execAsync(`osascript ${args}`, { timeout: timeoutMs });
  return stdout.trim();
}

/**
 * Set display options on the active or named window.
 * UNVERIFIED ON MAC — uses standard `display gridlines` / `display headings` /
 * `zoom` properties that have been in the Excel-for-Mac dictionary since 2011.
 */
export async function setDisplayOptionsViaAppleScript(
  filePath: string,
  sheetName: string | undefined,
  showGridlines: boolean | undefined,
  showRowColumnHeaders: boolean | undefined,
  zoomLevel: number | undefined,
  freezePaneCell: string | undefined,
  tabColor: string | undefined
): Promise<void> {
  const fileName = basename(filePath);
  const lines: string[] = [];
  lines.push(`tell application "Microsoft Excel"`);
  lines.push(`  set wb to workbook "${escapeAS(fileName)}"`);
  if (sheetName) {
    lines.push(`  set ws to worksheet "${escapeAS(sheetName)}" of wb`);
    lines.push(`  activate object ws`);
  }
  lines.push(`  set wnd to active window`);
  if (showGridlines !== undefined) {
    lines.push(`  set display gridlines of wnd to ${showGridlines}`);
  }
  if (showRowColumnHeaders !== undefined) {
    lines.push(`  set display headings of wnd to ${showRowColumnHeaders}`);
  }
  if (zoomLevel !== undefined) {
    lines.push(`  set zoom of wnd to ${Math.max(10, Math.min(400, zoomLevel))}`);
  }
  if (freezePaneCell !== undefined) {
    if (freezePaneCell === '') {
      lines.push(`  set freeze panes of wnd to false`);
    } else {
      lines.push(`  select range "${escapeAS(freezePaneCell)}"`);
      lines.push(`  set freeze panes of wnd to true`);
    }
  }
  if (tabColor !== undefined && sheetName) {
    if (tabColor === '') {
      lines.push(`  set color index of tab of ws to 0`); // xlColorIndexNone
    } else {
      // Parse #RRGGBB to {R, G, B} list (BGR-ish in Excel, but the property accepts a list)
      const hex = tabColor.replace('#', '');
      if (hex.length === 6) {
        const r = parseInt(hex.slice(0, 2), 16);
        const g = parseInt(hex.slice(2, 4), 16);
        const b = parseInt(hex.slice(4, 6), 16);
        lines.push(`  set color of tab of ws to {${r}, ${g}, ${b}}`);
      }
    }
  }
  lines.push(`  save wb`);
  lines.push(`end tell`);
  await runOsascript(lines.join('\n'));
}

/**
 * Protect or unprotect a sheet with optional password.
 * UNVERIFIED ON MAC — uses standard `protect` / `unprotect` commands.
 * Note: per-feature options (allowSort, allowFormatCells, etc.) are accepted
 * by Excel for Windows COM but not exposed via AppleScript on Mac. We honor
 * the password and protect/unprotect toggle only; the granular options are
 * silently dropped on Mac.
 */
export async function setSheetProtectionViaAppleScript(
  filePath: string,
  sheetName: string,
  protect: boolean,
  password: string | undefined
): Promise<void> {
  const fileName = basename(filePath);
  const lines: string[] = [];
  lines.push(`tell application "Microsoft Excel"`);
  lines.push(`  set ws to worksheet "${escapeAS(sheetName)}" of workbook "${escapeAS(fileName)}"`);
  if (protect) {
    if (password) {
      lines.push(`  protect ws password "${escapeAS(password)}"`);
    } else {
      lines.push(`  protect ws`);
    }
  } else {
    if (password) {
      lines.push(`  unprotect ws password "${escapeAS(password)}"`);
    } else {
      lines.push(`  unprotect ws`);
    }
  }
  lines.push(`  save workbook "${escapeAS(fileName)}"`);
  lines.push(`end tell`);
  await runOsascript(lines.join('\n'));
}

/**
 * Get the application-wide calculation mode.
 * UNVERIFIED ON MAC — uses standard `calculation` property which is app-wide
 * (not per-workbook) on Mac.
 */
export async function getCalculationModeViaAppleScript(_filePath: string): Promise<string> {
  const script = `tell application "Microsoft Excel" to get calculation`;
  const result = await runOsascript(script, 4000);
  // AppleScript returns one of: calculation automatic | calculation manual | calculation semiautomatic
  if (result.includes('manual')) return 'manual';
  if (result.includes('semi')) return 'semiautomatic';
  return 'automatic';
}

export async function setCalculationModeViaAppleScript(
  _filePath: string,
  mode: string
): Promise<void> {
  const m =
    mode === 'manual'
      ? 'calculation manual'
      : mode === 'semiautomatic'
      ? 'calculation semiautomatic'
      : 'calculation automatic';
  await runOsascript(`tell application "Microsoft Excel" to set calculation to ${m}`, 4000);
}

export async function triggerRecalculationViaAppleScript(
  _filePath: string,
  fullRecalc: boolean = false
): Promise<void> {
  const cmd = fullRecalc ? 'calculate full' : 'calculate';
  await runOsascript(`tell application "Microsoft Excel" to ${cmd}`, 8000);
}

/**
 * Capture a screenshot of the Excel window using the macOS `screencapture`
 * shell command. Less precise than the COM range-export-as-PNG on Windows
 * (this captures the whole window, not a specific range) but it works
 * without requiring AppleScript chart manipulation.
 */
export async function captureScreenshotViaAppleScript(
  filePath: string,
  _sheetName: string,
  outputPath: string,
  _range?: string
): Promise<void> {
  const fileName = basename(filePath);
  // Bring the window to front first so the screenshot has the right content
  await runOsascript(
    `tell application "Microsoft Excel" to activate\ntell application "Microsoft Excel" to activate window of workbook "${escapeAS(fileName)}"`,
    3000
  );
  // Tiny pause so the window is on top before capture
  await new Promise((r) => setTimeout(r, 250));
  // -x = no sound, -o = no shadow, -W = window mode would be interactive — instead
  // use AppleScript to get the window id and screencapture -l
  const idScript = `tell application "Microsoft Excel" to get id of window 1`;
  const winId = await runOsascript(idScript, 3000);
  await execAsync(`screencapture -x -o -l ${winId} "${outputPath.replace(/"/g, '\\"')}"`, {
    timeout: 8000,
  });
}

// ============================================================
// Live-mode INSPECTION tools (charts / pivot tables / shapes)
// ============================================================
// !!! UNVERIFIED ON MAC !!!
// All four functions below were authored from the Excel-for-Mac AppleScript
// dictionary docs but have NOT been validated on a real Mac. Property and
// command names (`chart objects`, `pivot tables`, `shapes`, `chart type`,
// `series collection`, `Pivot field type`, `pivot function`) appear in the
// Excel 2024 dictionary but the JSON-serialization shape of records over
// osascript is iffy and several attribute reads may throw. A coworker with a
// Mac needs to validate before this is shipped to non-Windows users.
//
// In particular we don't know:
//   - whether `chart type` returns the same numeric enum as Windows (xlChartType)
//   - whether `pivot function` exposes -4157 etc. directly
//   - how `formula` of a series serializes (likely a string, hopefully)
//   - whether RGB color values come back BGR-encoded or already RGB on Mac
//
// We DO emit JSON via `do shell script "echo ..."` to round-trip cleanly.
// Until the dictionary is validated, the dispatcher routes Mac to a clear
// error explaining this.

const UNVERIFIED_MAC_INSPECTION =
  'Live chart/pivot/shape inspection on Mac is unverified — the AppleScript implementation needs validation against a real Mac before being enabled. ' +
  'See docs/MAC_VERIFICATION_NEEDED.md for the validation checklist; for now use Windows + Excel COM, or open the file in Excel for Mac and read the values manually.';

/**
 * UNVERIFIED ON MAC. Stub for chart enumeration.
 * Real implementation should use:
 *   tell application "Microsoft Excel"
 *     repeat with sheet in worksheets of workbook X
 *       repeat with co in chart objects of sheet
 *         get name of co, chart type of chart of co, ...
 *       end repeat
 *     end repeat
 *   end tell
 */
export async function listChartsViaAppleScript(
  _filePath: string,
  _sheetName?: string
): Promise<string> {
  throw new Error(UNVERIFIED_MAC_INSPECTION);
}

/**
 * UNVERIFIED ON MAC. Stub for chart detail. See listChartsViaAppleScript.
 */
export async function getChartViaAppleScript(
  _filePath: string,
  _sheetName: string,
  _chartIndex?: number,
  _chartName?: string
): Promise<string> {
  throw new Error(UNVERIFIED_MAC_INSPECTION);
}

/**
 * UNVERIFIED ON MAC. Stub for pivot enumeration.
 * Real implementation should walk `pivot tables of sheet` and read
 * `name`, `source data`, `pivot fields`, and `pivot data fields` of each.
 */
export async function listPivotTablesViaAppleScript(
  _filePath: string,
  _sheetName?: string
): Promise<string> {
  throw new Error(UNVERIFIED_MAC_INSPECTION);
}

/**
 * UNVERIFIED ON MAC. Stub for shape enumeration.
 * Real implementation should walk `shapes of sheet` and read
 * `name`, `shape type`, `left position`, `top`, `width`, `height`,
 * `text frame`, and (if available) fill color.
 */
export async function listShapesViaAppleScript(
  _filePath: string,
  _sheetName?: string
): Promise<string> {
  throw new Error(UNVERIFIED_MAC_INSPECTION);
}
