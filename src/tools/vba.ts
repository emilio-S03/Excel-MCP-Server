import { isExcelRunningLive, isFileOpenInExcelLive } from './excel-live.js';
import { runVbaMacroLive, getVbaCodeLive, setVbaCodeLive, saveFileLive, checkVbaTrustLive, enableVbaTrustLive } from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

/** Patterns that create dialog boxes which freeze COM automation */
const INTERACTIVE_PATTERNS: Array<{ pattern: RegExp; name: string; description: string }> = [
  { pattern: /\bMsgBox\b/i, name: 'MsgBox', description: 'shows a popup dialog that requires a click to dismiss' },
  { pattern: /\bInputBox\b/i, name: 'InputBox', description: 'shows a text input dialog that waits for user typing' },
  { pattern: /\bApplication\.InputBox\b/i, name: 'Application.InputBox', description: 'shows an Excel input dialog that waits for user input' },
  { pattern: /\bApplication\.Dialogs\b/i, name: 'Application.Dialogs', description: 'opens a built-in Excel dialog window' },
  { pattern: /\bApplication\.FileDialog\b/i, name: 'Application.FileDialog', description: 'opens a file picker dialog' },
  { pattern: /\bApplication\.GetOpenFilename\b/i, name: 'Application.GetOpenFilename', description: 'opens a file-open dialog' },
  { pattern: /\bApplication\.GetSaveAsFilename\b/i, name: 'Application.GetSaveAsFilename', description: 'opens a save-as dialog' },
  { pattern: /\bUserForm\b/i, name: 'UserForm', description: 'displays a custom form window' },
  { pattern: /\b\.Show\b/i, name: '.Show', description: 'may display a form or dialog (review code carefully)' },
];

/** Patterns that don't freeze Excel but cause COM issues or mask errors */
const WARNING_PATTERNS: Array<{ pattern: RegExp; name: string; description: string }> = [
  { pattern: /\bOn\s+Error\s+Resume\s+Next\b/i, name: 'On Error Resume Next', description: 'suppresses all runtime errors — failures will be silent and invisible to the MCP error wrapper. If a command fails, execution continues with no indication. Consider using On Error GoTo with explicit handling instead.' },
  { pattern: /\bChr\s*\(\s*(\d+)\s*\)/i, name: 'Chr()', description: 'Chr() with values >255 crashes the COM connection and kills all VBA tools for the rest of the session. Use ASCII alternatives (e.g., "[!]" instead of Chr(9888)). Restarting Excel is the only recovery.' },
];

/**
 * Scan VBA code for interactive elements (MsgBox, InputBox, etc.)
 * that would freeze Excel when run via COM automation.
 * Returns list of warnings, empty if code is safe.
 */
function scanVbaForInteractiveElements(code: string): string[] {
  const warnings: string[] = [];
  const lines = code.split(/\r?\n/);

  for (const { pattern, name, description } of INTERACTIVE_PATTERNS) {
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      // Skip comment lines
      if (line.startsWith("'") || line.toLowerCase().startsWith('rem ')) continue;
      if (pattern.test(line)) {
        warnings.push(`Line ${i + 1}: "${name}" — ${description}`);
        break; // one warning per pattern is enough
      }
    }
  }

  return warnings;
}

/**
 * Scan VBA code for patterns that don't freeze Excel but cause COM issues
 * or mask errors (On Error Resume Next, Chr() with high values).
 * Returns list of warnings. These are non-blocking — code still runs.
 */
function scanVbaForWarnings(code: string): string[] {
  const warnings: string[] = [];
  const lines = code.split(/\r?\n/);

  for (const { pattern, name, description } of WARNING_PATTERNS) {
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      // Skip comment lines
      if (line.startsWith("'") || line.toLowerCase().startsWith('rem ')) continue;
      if (pattern.test(line)) {
        // Special handling for Chr(): only warn if value > 255
        if (name === 'Chr()') {
          const chrMatch = line.match(/\bChr\s*\(\s*(\d+)\s*\)/i);
          if (chrMatch && parseInt(chrMatch[1], 10) <= 255) continue;
        }
        warnings.push(`Line ${i + 1}: "${name}" — ${description}`);
        break; // one warning per pattern is enough
      }
    }
  }

  return warnings;
}

async function ensureFileOpenInExcel(filePath: string): Promise<void> {
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
}

export async function runVbaMacro(
  filePath: string,
  macroName: string,
  args: any[] = []
): Promise<string> {
  await ensureFileOpenInExcel(filePath);

  // Safety scan: try to read the macro's code and check for interactive elements + warnings
  let interactiveWarnings: string[] = [];
  let comWarnings: string[] = [];
  try {
    // macroName could be "Module1.MyMacro" — extract the module name
    const moduleName = macroName.includes('.') ? macroName.split('.')[0] : 'Module1';
    const code = await getVbaCodeLive(filePath, moduleName);
    if (code) {
      interactiveWarnings = scanVbaForInteractiveElements(code);
      comWarnings = scanVbaForWarnings(code);
    }
  } catch {
    // If we can't read the code (trust issue, wrong module name), skip the scan
  }

  if (interactiveWarnings.length > 0) {
    return JSON.stringify({
      success: false,
      blocked: true,
      message: `Macro "${macroName}" was NOT executed because it contains interactive elements that would freeze Excel.`,
      interactiveElements: interactiveWarnings,
      method: 'live',
      suggestion: 'Remove MsgBox/InputBox calls from the macro, or replace them with writing results to cells. Then try again.',
    }, null, 2);
  }

  try {
    const result = await runVbaMacroLive(filePath, macroName, args);

    const response: Record<string, any> = {
      success: true,
      message: `Macro "${macroName}" executed`,
      result: result || null,
      method: 'live',
      note: 'VBA macros require "Trust access to VBA project object model" to be enabled in Excel Trust Center.',
    };
    if (comWarnings.length > 0) {
      response.comWarnings = comWarnings;
      response.warningNote = 'This macro contains patterns that may cause issues with COM automation. If the macro appears to succeed but nothing happens, these warnings may explain why.';
    }
    return JSON.stringify(response, null, 2);
  } catch (error: any) {
    const msg = error.message || String(error);
    // Detect wrapped VBA runtime errors from the error handler injection
    const vbaMatch = msg.match(/VBA_RUNTIME_ERROR\|(\d+)\|([^|]*)\|?(.*)?/);
    if (vbaMatch) {
      return JSON.stringify({
        success: false,
        vbaError: true,
        message: `VBA runtime error in macro "${macroName}"`,
        errorNumber: parseInt(vbaMatch[1], 10),
        errorDescription: vbaMatch[2].trim(),
        errorSource: vbaMatch[3]?.trim() || '',
        method: 'live',
        suggestion: 'Fix the VBA code error and try again. The error was caught by the MCP error handler — Excel did NOT freeze.',
      }, null, 2);
    }
    // Re-throw non-VBA errors
    throw error;
  }
}

export async function getVbaCode(
  filePath: string,
  moduleName: string
): Promise<string> {
  await ensureFileOpenInExcel(filePath);
  const code = await getVbaCodeLive(filePath, moduleName);

  return JSON.stringify({
    moduleName,
    code: code || '',
    method: 'live',
    note: 'VBA access requires "Trust access to VBA project object model" to be enabled in Excel Trust Center.',
  }, null, 2);
}

export async function setVbaCode(
  filePath: string,
  moduleName: string,
  code: string,
  createModule: boolean = false,
  appendMode: boolean = false
): Promise<string> {
  await ensureFileOpenInExcel(filePath);

  // Safety scan: warn about interactive elements and COM-unsafe patterns
  const interactiveWarnings = scanVbaForInteractiveElements(code);
  const comWarnings = scanVbaForWarnings(code);

  await setVbaCodeLive(filePath, moduleName, code, createModule, appendMode);
  await saveFileLive(filePath);

  const response: Record<string, any> = {
    success: true,
    message: `VBA code ${appendMode ? 'appended to' : createModule ? 'created in' : 'updated in'} module "${moduleName}"`,
    moduleName,
    method: 'live',
    note: 'VBA access requires "Trust access to VBA project object model" to be enabled in Excel Trust Center.',
  };

  if (interactiveWarnings.length > 0) {
    response.warning = 'This code contains interactive elements (MsgBox, InputBox, etc.) that will FREEZE Excel if run via COM automation. The code was saved, but DO NOT run this macro via excel_run_vba_macro — it will hang.';
    response.interactiveElements = interactiveWarnings;
    response.suggestion = 'Replace MsgBox/InputBox with writing results to cells, or use Debug.Print instead.';
  }

  if (comWarnings.length > 0) {
    response.comWarnings = comWarnings;
    response.comWarningNote = 'This code contains patterns that may cause issues with COM automation. Review the warnings above before running this macro.';
  }

  return JSON.stringify(response, null, 2);
}

export async function checkVbaTrust(): Promise<string> {
  const raw = await checkVbaTrustLive();
  const result = JSON.parse(raw);

  let explanation: string;
  if (result.ComTestPassed) {
    explanation = 'VBA trust is WORKING. COM can access VBE. VBA tools will work with .xlsm files.';
  } else if (result.Enabled) {
    explanation = 'Registry says ENABLED but COM test FAILED. This means Excel is not reading from the registry paths we found. FIX: Manually check the box in Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to the VBA project object model". Then restart Excel fully (Task Manager → end all EXCEL.EXE).';
  } else {
    explanation = 'VBA trust is DISABLED. To fix: (1) Use excel_enable_vba_trust to set registry keys, OR (2) manually check the box in Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to the VBA project object model". Then restart Excel completely.';
  }

  return JSON.stringify({
    ...result,
    method: 'registry+com',
    explanation,
  }, null, 2);
}

export async function enableVbaTrust(
  enable: boolean
): Promise<string> {
  const raw = await enableVbaTrustLive(enable);
  const result = JSON.parse(raw);

  if (!result.Success) {
    throw new Error(result.Note || 'Failed to change VBA trust setting');
  }

  return JSON.stringify({
    ...result,
    method: 'registry',
    explanation: enable
      ? 'Registry keys written to ALL found Office paths. You MUST restart Excel completely (close ALL windows, verify no EXCEL.EXE in Task Manager, then reopen). If VBA still does not work after restart, the manual checkbox may be needed: Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → check "Trust access to the VBA project object model".'
      : 'VBA trust access is now DISABLED. Restart Excel for the change to take effect.',
    important: 'This changes a security setting. VBA trust allows programmatic access to VBA code in Excel workbooks. Only enable this if you trust the workbooks you are opening.',
  }, null, 2);
}
