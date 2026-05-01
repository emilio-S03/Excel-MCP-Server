/**
 * excel_check_environment — startup capability probe.
 *
 * Returns a structured report so callers (Claude, or human users) know what
 * the server can actually do on this machine BEFORE invoking failing tools.
 * Avoids first-invocation hangs from un-granted Mac Automation permission
 * and surfaces VBA trust state without making the user discover it through
 * a cryptic COM error.
 */
import { exec } from 'child_process';
import { promisify } from 'util';
import { platform } from 'os';
import { getAllowedDirectories } from './helpers.js';

const execAsync = promisify(exec);
const IS_WIN = platform() === 'win32';
const IS_MAC = platform() === 'darwin';
const IS_LINUX = platform() === 'linux';

export interface EnvironmentReport {
  platform: 'win32' | 'darwin' | 'linux';
  nodeVersion: string;
  serverVersion: string;
  excel: {
    installed: boolean | 'unknown';
    running: boolean;
    version: string | null;
    detectionMethod: string;
  };
  vbaTrust?: {
    enabled: boolean | 'unknown';
    note: string;
  };
  automationPermission?: {
    granted: boolean | 'unknown';
    note: string;
  };
  capabilityMatrix: {
    fileMode: { available: boolean; tools: string[] };
    liveMode: { available: boolean; reason: string | null };
    vbaTools: { available: boolean; reason: string | null };
    powerQuery: { available: boolean; reason: string | null };
  };
  config: {
    allowedDirectories: string[];
    allowedDirectoriesIsDefault: boolean;
  };
  recommendations: string[];
}

async function runWithTimeout(cmd: string, timeoutMs: number): Promise<string> {
  const { stdout } = await execAsync(cmd, { timeout: timeoutMs });
  return stdout.trim();
}

async function detectExcelWindows(): Promise<{ running: boolean; installed: boolean | 'unknown'; version: string | null; detection: string }> {
  let running = false;
  let installed: boolean | 'unknown' = 'unknown';
  let version: string | null = null;
  try {
    const out = await runWithTimeout(
      'powershell.exe -NoProfile -NonInteractive -Command "if (Get-Process EXCEL -ErrorAction SilentlyContinue) { Write-Output running } else { Write-Output not-running }"',
      5000
    );
    running = out === 'running';
  } catch {
    /* ignore */
  }

  try {
    const out = await runWithTimeout(
      'powershell.exe -NoProfile -NonInteractive -Command "$p = Get-Command excel.exe -ErrorAction SilentlyContinue; if ($p) { (Get-Item $p.Source).VersionInfo.ProductVersion } else { Write-Output __NOT_FOUND__ }"',
      5000
    );
    if (out && out !== '__NOT_FOUND__') {
      installed = true;
      version = out;
    } else if (running) {
      installed = true; // it's running, so it must be installed even if not on PATH
    } else {
      installed = 'unknown';
    }
  } catch {
    /* ignore */
  }

  return { running, installed, version, detection: 'Get-Process + Get-Command excel.exe' };
}

async function detectExcelMac(): Promise<{ running: boolean; installed: boolean | 'unknown'; version: string | null; detection: string }> {
  let running = false;
  let installed: boolean | 'unknown' = 'unknown';
  let version: string | null = null;
  try {
    const out = await runWithTimeout(
      `osascript -e 'tell application "System Events" to (name of processes) contains "Microsoft Excel"'`,
      3000
    );
    running = out === 'true';
  } catch {
    /* ignore */
  }

  try {
    const out = await runWithTimeout(
      `defaults read /Applications/Microsoft\\ Excel.app/Contents/Info.plist CFBundleShortVersionString`,
      3000
    );
    if (out) {
      installed = true;
      version = out;
    }
  } catch {
    installed = 'unknown';
  }

  return { running, installed, version, detection: 'osascript + Info.plist' };
}

async function checkVbaTrustWindows(): Promise<{ enabled: boolean | 'unknown'; note: string }> {
  // Probe several Excel versions; first hit wins.
  const candidatePaths = [
    'HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Security',
    'HKCU:\\Software\\Microsoft\\Office\\15.0\\Excel\\Security',
    'HKCU:\\Software\\Microsoft\\Office\\14.0\\Excel\\Security',
  ];
  for (const path of candidatePaths) {
    try {
      const out = await runWithTimeout(
        `powershell.exe -NoProfile -NonInteractive -Command "(Get-ItemProperty -Path '${path}' -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM"`,
        4000
      );
      if (out === '1') {
        return { enabled: true, note: `Trusted via ${path}` };
      }
      if (out === '0') {
        return {
          enabled: false,
          note: `AccessVBOM=0 at ${path}. To enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > "Trust access to the VBA project object model".`,
        };
      }
    } catch {
      /* try next */
    }
  }
  return {
    enabled: 'unknown',
    note: 'Could not read VBA trust registry. Excel may not be installed, or the version is older than 14.0 (2010).',
  };
}

async function checkAutomationPermissionMac(): Promise<{ granted: boolean | 'unknown'; note: string }> {
  // The only reliable way to detect this is to *try* an osascript and see if
  // it hangs/fails. We use a 2s timeout — if it returns "Microsoft Excel" we
  // have permission, if it errors with "not allowed" we don't, if it times
  // out we don't either.
  try {
    const out = await runWithTimeout(
      `osascript -e 'tell application "Microsoft Excel" to get name'`,
      2500
    );
    if (out.includes('Microsoft Excel') || out.includes('Excel')) {
      return { granted: true, note: 'Verified via test osascript call' };
    }
    return { granted: 'unknown', note: `Unexpected response: ${out}` };
  } catch (err: any) {
    const msg = String(err?.stderr || err?.message || '');
    if (msg.includes('not allowed') || msg.includes('-1743')) {
      return {
        granted: false,
        note: 'macOS Automation permission denied. Grant access: System Settings > Privacy & Security > Automation > Claude (Desktop or Code) > enable Microsoft Excel.',
      };
    }
    if (err?.killed) {
      return {
        granted: 'unknown',
        note: 'osascript timed out — likely a permission prompt is hidden behind another window. Switch to Microsoft Excel and try a tool to surface the prompt.',
      };
    }
    return { granted: 'unknown', note: msg || 'Unknown osascript failure' };
  }
}

const FILE_MODE_TOOLS = [
  'excel_read_workbook', 'excel_read_sheet', 'excel_read_range',
  'excel_get_cell', 'excel_get_formula', 'excel_write_workbook',
  'excel_update_cell', 'excel_write_range', 'excel_add_row', 'excel_set_formula',
  'excel_format_cell', 'excel_set_column_width', 'excel_set_row_height', 'excel_merge_cells',
  'excel_create_sheet', 'excel_delete_sheet', 'excel_rename_sheet', 'excel_duplicate_sheet',
  'excel_delete_rows', 'excel_delete_columns', 'excel_copy_range',
  'excel_search_value', 'excel_filter_rows',
  'excel_create_chart', 'excel_create_pivot_table', 'excel_create_table',
  'excel_validate_formula_syntax', 'excel_validate_range', 'excel_get_data_validation_info',
  'excel_insert_rows', 'excel_insert_columns', 'excel_unmerge_cells', 'excel_get_merged_cells',
  'excel_apply_conditional_format',
  'excel_add_image', 'excel_csv_import', 'excel_csv_export', 'excel_find_replace',
  // v3.1
  'excel_get_conditional_formats', 'excel_list_data_validations', 'excel_get_sheet_protection',
  'excel_get_display_options', 'excel_get_workbook_properties', 'excel_set_workbook_properties',
  'excel_get_hyperlinks', 'excel_sort', 'excel_set_auto_filter', 'excel_clear_auto_filter',
  'excel_remove_duplicates', 'excel_paste_special',
  'excel_set_sheet_visibility', 'excel_list_sheet_visibility', 'excel_hide_rows', 'excel_hide_columns',
  'excel_add_hyperlink', 'excel_remove_hyperlink',
  'excel_add_sparkline', 'excel_remove_sparklines',
  'excel_get_page_setup', 'excel_set_page_setup',
  // formula auditing & workbook health-check
  'excel_find_formula_errors', 'excel_find_circular_references', 'excel_workbook_stats',
  'excel_list_formulas', 'excel_trace_precedents',
];

export async function checkEnvironment(): Promise<string> {
  const plat = (platform() === 'win32' ? 'win32' : platform() === 'darwin' ? 'darwin' : 'linux') as
    | 'win32' | 'darwin' | 'linux';

  const recommendations: string[] = [];

  let excel: EnvironmentReport['excel'];
  let vbaTrust: EnvironmentReport['vbaTrust'] | undefined;
  let automationPermission: EnvironmentReport['automationPermission'] | undefined;

  if (IS_WIN) {
    const r = await detectExcelWindows();
    excel = { installed: r.installed, running: r.running, version: r.version, detectionMethod: r.detection };
    vbaTrust = await checkVbaTrustWindows();
    if (vbaTrust.enabled === false) recommendations.push(vbaTrust.note);
    if (excel.installed === 'unknown') recommendations.push('Could not detect Excel on PATH. If installed, the file-mode tools (~38 tools) still work; live-mode (COM) tools require it to be running.');
  } else if (IS_MAC) {
    const r = await detectExcelMac();
    excel = { installed: r.installed, running: r.running, version: r.version, detectionMethod: r.detection };
    automationPermission = await checkAutomationPermissionMac();
    if (automationPermission.granted === false) recommendations.push(automationPermission.note);
    if (excel.installed !== true) recommendations.push('Microsoft Excel not detected at /Applications/Microsoft Excel.app. File-mode tools work without it; live-mode tools need it.');
  } else {
    excel = { installed: false, running: false, version: null, detectionMethod: 'linux: no Excel detection' };
    recommendations.push('On Linux, only file-mode tools (no live Excel automation) are available.');
  }

  const liveAvailable = excel.running && (
    (IS_WIN) ||
    (IS_MAC && automationPermission?.granted === true)
  );
  const vbaAvailable = IS_WIN && excel.running && vbaTrust?.enabled === true;
  const powerQueryAvailable = IS_WIN && excel.running;

  const { dirs, isDefault } = getAllowedDirectories();

  if (isDefault) {
    recommendations.push('Sandbox is using defaults (Documents/Desktop/Downloads). To allow other folders, set EXCEL_ALLOWED_DIRS in your MCP server config.');
  }

  const report: EnvironmentReport = {
    platform: plat,
    nodeVersion: process.version,
    serverVersion: "3.2.0",
    excel,
    ...(vbaTrust ? { vbaTrust } : {}),
    ...(automationPermission ? { automationPermission } : {}),
    capabilityMatrix: {
      fileMode: { available: true, tools: FILE_MODE_TOOLS },
      liveMode: {
        available: liveAvailable,
        reason: liveAvailable ? null :
          !excel.running ? 'Excel is not running. Open the file in Excel for live editing.' :
          IS_MAC && automationPermission?.granted !== true ? 'macOS Automation permission not granted.' :
          'Unknown.',
      },
      vbaTools: {
        available: vbaAvailable,
        reason: vbaAvailable ? null :
          IS_MAC ? 'VBA execution via AppleScript was disabled by Microsoft on Excel for Mac. Use Office Scripts (cloud) instead.' :
          IS_LINUX ? 'No Excel on Linux — VBA tools unavailable.' :
          !excel.running ? 'Excel must be running.' :
          vbaTrust?.enabled !== true ? (vbaTrust?.note || 'VBA Trust Center access not enabled.') :
          'Unknown.',
      },
      powerQuery: {
        available: powerQueryAvailable,
        reason: powerQueryAvailable ? null :
          IS_MAC ? 'Power Query automation not supported on Mac.' :
          IS_LINUX ? 'No Excel on Linux.' :
          'Excel must be running.',
      },
    },
    config: { allowedDirectories: dirs, allowedDirectoriesIsDefault: isDefault },
    recommendations,
  };

  return JSON.stringify(report, null, 2);
}
