import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import { resolve, dirname } from 'path';
import { homedir } from 'os';
import { ERROR_MESSAGES } from '../constants.js';
import type { CellValue } from 'exceljs';

// ----------------------------------------------------------------------------
// User configuration (Phase 1 — v3.0.0)
// Hydrated once at boot from environment variables. The .mcpb spec uses
// ${user_config.NAME} substitution to inject these into the spawned process.
// Manual installs (raw claude_desktop_config.json / ~/.claude.json) set the
// same env vars in their server entry. There is no runtime config channel
// — Spike B confirmed Claude Desktop and Claude Code don't deliver one.
// ----------------------------------------------------------------------------

function defaultAllowedDirectories(): string[] {
  const home = homedir();
  return [
    resolve(home, 'Documents'),
    resolve(home, 'Desktop'),
    resolve(home, 'Downloads'),
  ];
}

function parseDirList(raw: string | undefined): string[] | null {
  if (!raw) return null;
  // Accept either OS-specific path separator (':' on POSIX, ';' on win32)
  // or a JSON array. Trim and drop empties.
  const trimmed = raw.trim();
  if (trimmed.startsWith('[')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed.map(String).map((p) => resolve(p));
    } catch {
      // fall through
    }
  }
  const sep = process.platform === 'win32' ? /[;,]/ : /[:;,]/;
  const parts = trimmed.split(sep).map((s) => s.trim()).filter(Boolean);
  return parts.length > 0 ? parts.map((p) => resolve(p)) : null;
}

let allowedDirectories: string[] = [];
let allowedDirectoriesIsDefault = true;

function loadAllowedDirsFromEnv(): void {
  const fromEnv = parseDirList(process.env.EXCEL_ALLOWED_DIRS);
  if (fromEnv && fromEnv.length > 0) {
    allowedDirectories = fromEnv;
    allowedDirectoriesIsDefault = false;
  } else {
    allowedDirectories = defaultAllowedDirectories();
    allowedDirectoriesIsDefault = true;
  }
}

loadAllowedDirsFromEnv();

export function setAllowedDirectories(dirs: string[]) {
  if (!Array.isArray(dirs) || dirs.length === 0) {
    allowedDirectories = defaultAllowedDirectories();
    allowedDirectoriesIsDefault = true;
    return;
  }
  allowedDirectories = dirs.map((d) => resolve(d));
  allowedDirectoriesIsDefault = false;
}

export function getAllowedDirectories(): { dirs: string[]; isDefault: boolean } {
  return { dirs: [...allowedDirectories], isDefault: allowedDirectoriesIsDefault };
}

export function ensureFilePathAllowed(filePath: string): void {
  // Empty allow-list should never happen now (defaults always populate),
  // but keep the safety check in case future code clears it.
  if (!allowedDirectories || allowedDirectories.length === 0) {
    return;
  }

  const absolutePath = resolve(filePath);
  const pathDir = dirname(absolutePath);

  const isAllowed = allowedDirectories.some((allowedDir) => {
    const absoluteAllowedDir = resolve(allowedDir);
    return (
      absolutePath === absoluteAllowedDir ||
      absolutePath.startsWith(absoluteAllowedDir + (process.platform === 'win32' ? '\\' : '/')) ||
      pathDir === absoluteAllowedDir ||
      pathDir.startsWith(absoluteAllowedDir + (process.platform === 'win32' ? '\\' : '/'))
    );
  });

  if (!isAllowed) {
    const err = new Error(
      `PATH_OUTSIDE_ALLOWED: "${filePath}" is not within an allowed directory. ` +
      `Allowed: ${allowedDirectories.join(', ')}. ` +
      `To allow more locations, set the EXCEL_ALLOWED_DIRS environment variable in your MCP server config ` +
      `(e.g., "C:/Users/you/Projects;C:/Data" on Windows, "/Users/you/Projects:/Data" on macOS).`
    );
    (err as any).code = 'PATH_OUTSIDE_ALLOWED';
    throw err;
  }
}

export async function loadWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
  // Validate file path against allowed directories
  ensureFilePathAllowed(filePath);

  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`${ERROR_MESSAGES.FILE_NOT_FOUND}: ${filePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
    return workbook;
  } catch (error) {
    throw new Error(`${ERROR_MESSAGES.READ_ERROR}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

export function getSheet(workbook: ExcelJS.Workbook, sheetName: string): ExcelJS.Worksheet {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`${ERROR_MESSAGES.SHEET_NOT_FOUND}: ${sheetName}`);
  }
  return sheet;
}

export async function saveWorkbook(workbook: ExcelJS.Workbook, filePath: string, createBackup: boolean = false): Promise<void> {
  // Validate file path against allowed directories
  ensureFilePathAllowed(filePath);

  try {
    if (createBackup) {
      try {
        await fs.access(filePath);
        const backupPath = `${filePath}.backup`;
        await fs.copyFile(filePath, backupPath);
      } catch {
        // File doesn't exist, no backup needed
      }
    }

    await workbook.xlsx.writeFile(filePath);
  } catch (error) {
    throw new Error(`${ERROR_MESSAGES.WRITE_ERROR}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

export function columnLetterToNumber(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

export function columnNumberToLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

export function cellValueToString(value: CellValue): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'object') {
    if ('formula' in value && value.formula) {
      return `=${value.formula}`;
    }
    if ('result' in value) {
      return String(value.result);
    }
    if ('text' in value) {
      return String(value.text);
    }
    return JSON.stringify(value);
  }
  return String(value);
}

export function formatDataAsTable(data: any[][], headers?: string[]): string {
  if (data.length === 0) {
    return 'No data';
  }

  const tableData = headers ? [headers, ...data] : data;
  const colWidths: number[] = [];

  // Calculate column widths
  for (const row of tableData) {
    for (let i = 0; i < row.length; i++) {
      const cellText = cellValueToString(row[i]);
      colWidths[i] = Math.max(colWidths[i] || 0, cellText.length);
    }
  }

  // Build table
  let table = '';
  for (let i = 0; i < tableData.length; i++) {
    const row = tableData[i];
    const cells = row.map((cell, j) => {
      const text = cellValueToString(cell);
      return text.padEnd(colWidths[j] || 0);
    });
    table += '| ' + cells.join(' | ') + ' |\n';

    // Add separator after header
    if (headers && i === 0) {
      table += '| ' + colWidths.map(w => '-'.repeat(w)).join(' | ') + ' |\n';
    }
  }

  return table;
}

export function parseRange(range: string): { startCol: number; startRow: number; endCol: number; endRow: number } {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(ERROR_MESSAGES.INVALID_RANGE);
  }

  return {
    startCol: columnLetterToNumber(match[1]),
    startRow: parseInt(match[2]),
    endCol: columnLetterToNumber(match[3]),
    endRow: parseInt(match[4]),
  };
}
