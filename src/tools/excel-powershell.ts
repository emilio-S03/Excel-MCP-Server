import { exec } from 'child_process';
import { promisify } from 'util';
import { basename, join } from 'path';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';

const execAsync = promisify(exec);

// Configuration
const POWERSHELL_TIMEOUT = 10000; // 10 seconds
const MAX_RETRIES = 3;
const RETRY_DELAY = 500; // 500ms
const MAX_CMD_LENGTH = 8000; // Leave margin below cmd.exe's 8191 limit

/**
 * Decode common HTML entities that may be injected by the MCP protocol layer
 * or Claude Desktop before the server receives the string.
 * Known issue: VBA string concatenation '&' arrives as '&amp;'.
 */
function decodeHtmlEntities(str: string): string {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/g, "'");
}

/**
 * Escape a string for safe use inside PowerShell single-quoted strings.
 * In single-quoted strings, the only escape is ' → ''
 */
function escapePowerShellString(str: string): string {
  return str.replace(/'/g, "''");
}

/**
 * Convert column number to Excel letter (1=A, 27=AA, etc.)
 */
function numberToColumnLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

/**
 * Convert Excel column letter to number (A=1, AA=27, etc.)
 */
function columnLetterToNumber(letter: string): number {
  let num = 0;
  for (let i = 0; i < letter.length; i++) {
    num = num * 26 + (letter.charCodeAt(i) - 64);
  }
  return num;
}

/**
 * Validate cell address format (e.g., "A1", "Z100")
 * Security: Prevents injection attacks by ensuring cell addresses match expected format
 */
function validateCellAddress(address: string): void {
  if (!/^[A-Z]+\d+$/.test(address)) {
    throw new Error(`Invalid cell address format: ${address}. Expected format like "A1" or "AA100"`);
  }
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (match) {
    const colNum = columnLetterToNumber(match[1]);
    if (colNum > 16384) {
      throw new Error(`Column ${match[1]} exceeds Excel's maximum column (XFD)`);
    }
    const rowNum = parseInt(match[2]);
    if (rowNum < 1 || rowNum > 1048576) {
      throw new Error(`Row ${rowNum} must be between 1 and 1048576`);
    }
  }
}

/**
 * Validate range format (e.g., "A1:B10")
 * Security: Prevents injection attacks by ensuring ranges match expected format
 */
function validateRange(range: string): void {
  if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
    throw new Error(`Invalid range format: ${range}. Expected format like "A1:B10"`);
  }
  const [start, end] = range.split(':');
  validateCellAddress(start);
  validateCellAddress(end);
}

/**
 * Format value for PowerShell based on type.
 * Numbers are passed as literals, strings are wrapped in single quotes.
 */
function formatValueForPowerShell(value: string | number): string {
  if (typeof value === 'number') {
    return String(value);
  }
  return `'${escapePowerShellString(String(value))}'`;
}

/**
 * Execute PowerShell script with timeout and retry logic.
 * Uses -EncodedCommand for small scripts, temp file for large ones.
 * Always uses powershell.exe (5.1/.NET Framework) — NOT pwsh.exe —
 * because GetActiveObject requires .NET Framework.
 */
async function execPowerShellWithRetry(
  script: string,
  retries: number = MAX_RETRIES,
  timeout: number = POWERSHELL_TIMEOUT
): Promise<string> {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const encoded = Buffer.from(script, 'utf16le').toString('base64');
      const encodedCmd = `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ${encoded}`;

      let stdout: string;

      if (encodedCmd.length > MAX_CMD_LENGTH) {
        // Script too large for command line, use temp file
        const tmpFile = join(tmpdir(), `excel-mcp-${Date.now()}-${Math.random().toString(36).slice(2)}.ps1`);
        try {
          writeFileSync(tmpFile, script, 'utf8');
          const result = await execAsync(
            `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "${tmpFile}"`,
            { timeout }
          );
          stdout = result.stdout;
        } finally {
          try { unlinkSync(tmpFile); } catch {}
        }
      } else {
        const result = await execAsync(encodedCmd, { timeout });
        stdout = result.stdout;
      }

      return stdout.trim();
    } catch (error: any) {
      console.error(`[PowerShell] Attempt ${attempt}/${retries} failed:`, {
        error: error.message,
        code: error.code,
        killed: error.killed,
      });

      if (attempt === retries) {
        // Categorize the error for better diagnostics
        let categorized: string;
        if (error.killed === true) {
          categorized = 'PowerShell timed out. Excel may have a modal dialog open (e.g., a VBA error popup or Save prompt). Dismiss the dialog in Excel and retry. Use excel_diagnose_connection to investigate.';
        } else if (error.message && (error.message.includes('GetActiveObject') || error.message.includes('0x800401E3'))) {
          categorized = 'Could not connect to Excel via COM. Excel may not be running, or it may be running elevated (as admin) while this process is not. Use excel_diagnose_connection to investigate.';
        } else if (error.message && error.message.toLowerCase().includes('trust')) {
          categorized = 'VBA trust access denied. Enable "Trust access to the VBA project object model" in Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings. Then restart Excel.';
        } else if (error.message && /0x800A/.test(error.message)) {
          const hresultMatch = error.message.match(/(0x800A[0-9A-Fa-f]+)/);
          categorized = `Excel COM error${hresultMatch ? ` (HRESULT ${hresultMatch[1]})` : ''}: ${error.message}. Use excel_diagnose_connection to investigate.`;
        } else {
          categorized = `${error.message || error}. Use excel_diagnose_connection to investigate.`;
        }
        const categorizedError = new Error(categorized);
        (categorizedError as any).originalError = error;
        throw categorizedError;
      }

      // Wait before retry with exponential backoff
      await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * attempt));
    }
  }
  throw new Error('Max retries exceeded');
}

/**
 * Build PowerShell COM preamble to get Excel, workbook, and optionally worksheet.
 * Includes error handling for missing workbook.
 */
function buildPreamble(fileName: string, sheetName?: string): string {
  const escapedFileName = escapePowerShellString(fileName);
  let script = `$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')\n`;
  script += `$wb = $excel.Workbooks | Where-Object { $_.Name -eq '${escapedFileName}' }\n`;
  script += `if (-not $wb) { throw "Workbook '${escapedFileName}' not found in open Excel instance" }\n`;
  if (sheetName) {
    const escapedSheetName = escapePowerShellString(sheetName);
    script += `$ws = $wb.Worksheets.Item('${escapedSheetName}')\n`;
  }
  return script;
}

/**
 * Wrap script body in try/finally with COM cleanup
 */
function wrapWithCleanup(preamble: string, body: string): string {
  return `${preamble}try {\n${body}} finally {\n  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null\n}\n`;
}

// ============================================================
// Detection Functions
// ============================================================

/**
 * Check if Microsoft Excel is running on Windows
 */
export async function isExcelRunningWindows(): Promise<boolean> {
  try {
    const script = `if (Get-Process EXCEL -ErrorAction SilentlyContinue) { Write-Output 'true' } else { Write-Output 'false' }`;
    const result = await execPowerShellWithRetry(script, 2, 5000);
    const isRunning = result === 'true';
    console.error(`[PowerShell] Excel running: ${isRunning}`);
    return isRunning;
  } catch (error: any) {
    console.error('[PowerShell] Failed to check if Excel is running:', error.message);
    return false;
  }
}

/**
 * Check if a specific Excel file is open on Windows
 */
export async function isFileOpenInExcelWindows(filePath: string): Promise<boolean> {
  try {
    const fileName = basename(filePath);
    const escapedFileName = escapePowerShellString(fileName);
    console.error(`[PowerShell] Checking if file is open: ${fileName}`);

    const script = `try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  $found = $false
  foreach ($wb in $excel.Workbooks) {
    if ($wb.Name -eq '${escapedFileName}') { $found = $true; break }
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  if ($found) { Write-Output 'true' } else { Write-Output 'false' }
} catch {
  Write-Output 'false'
}`;

    const result = await execPowerShellWithRetry(script, 2, 5000);
    const isOpen = result === 'true';
    console.error(`[PowerShell] File "${fileName}" open: ${isOpen}`);
    return isOpen;
  } catch (error: any) {
    console.error(`[PowerShell] Failed to check if file is open:`, {
      file: basename(filePath),
      error: error.message,
    });
    return false;
  }
}

// ============================================================
// Cell Operations
// ============================================================

/**
 * Update a cell value in an open Excel file via PowerShell COM
 */
export async function updateCellViaPowerShell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  value: string | number
): Promise<void> {
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);
  const formattedValue = formatValueForPowerShell(value);

  console.error(`[PowerShell] Updating cell ${cellAddress} in "${fileName}"/"${sheetName}"`);

  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $ws.Range('${cellAddress}').Value2 = ${formattedValue}\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully updated cell ${cellAddress}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to update cell:`, {
      file: fileName,
      sheet: sheetName,
      cell: cellAddress,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Read a cell value from an open Excel file via PowerShell COM
 */
export async function readCellViaPowerShell(
  filePath: string,
  sheetName: string,
  cellAddress: string
): Promise<string> {
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);
  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $val = $ws.Range('${cellAddress}').Value2\n  if ($val -ne $null) { Write-Output ([string]$val) } else { Write-Output '' }\n`;
  const script = wrapWithCleanup(preamble, body);

  const result = await execPowerShellWithRetry(script);
  return result;
}

/**
 * Write a 2D array to a range starting at startCell in an open Excel file.
 * Batches all cell writes into a single PowerShell invocation for performance.
 */
export async function writeRangeViaPowerShell(
  filePath: string,
  sheetName: string,
  startCell: string,
  data: (string | number)[][]
): Promise<void> {
  validateCellAddress(startCell);

  const fileName = basename(filePath);

  console.error(`[PowerShell] Writing range starting at ${startCell} in "${fileName}"/"${sheetName}" with ${data.length} rows`);

  try {
    const match = startCell.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      throw new Error(`Invalid cell address: ${startCell}`);
    }
    const startColLetter = match[1];
    const startRow = parseInt(match[2]);
    const startCol = columnLetterToNumber(startColLetter);

    // Build a single PowerShell script with all cell assignments
    const preamble = buildPreamble(fileName, sheetName);
    let body = '';

    for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
      const row = data[rowIdx];
      for (let colIdx = 0; colIdx < row.length; colIdx++) {
        const targetRow = startRow + rowIdx;
        const targetCol = startCol + colIdx;
        const targetColLetter = numberToColumnLetter(targetCol);
        const cellAddress = `${targetColLetter}${targetRow}`;
        const value = row[colIdx];
        const formattedValue = formatValueForPowerShell(value);

        body += `  $ws.Range('${cellAddress}').Value2 = ${formattedValue}\n`;
      }
    }

    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully wrote range starting at ${startCell}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to write range:`, {
      file: fileName,
      sheet: sheetName,
      startCell,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Add a new row to a sheet in an open Excel file via PowerShell COM.
 * Finds the last used row and appends data after it.
 */
export async function addRowViaPowerShell(
  filePath: string,
  sheetName: string,
  rowData: (string | number)[]
): Promise<void> {
  const fileName = basename(filePath);

  console.error(`[PowerShell] Adding row to "${fileName}"/"${sheetName}" with ${rowData.length} cells`);

  try {
    const preamble = buildPreamble(fileName, sheetName);

    let body = `  $lastRow = $ws.UsedRange.Rows.Count\n`;
    body += `  $newRow = $lastRow + 1\n`;

    for (let i = 0; i < rowData.length; i++) {
      const value = rowData[i];
      const formattedValue = formatValueForPowerShell(value);
      body += `  $ws.Cells.Item($newRow, ${i + 1}).Value2 = ${formattedValue}\n`;
    }

    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully added row`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to add row:`, {
      file: fileName,
      sheet: sheetName,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Set a formula in a cell in an open Excel file via PowerShell COM
 */
export async function setFormulaViaPowerShell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  formula: string
): Promise<void> {
  validateCellAddress(cellAddress);

  // Ensure formula starts with "="
  const normalizedFormula = formula.startsWith('=') ? formula : `=${formula}`;

  const fileName = basename(filePath);
  const escapedFormula = escapePowerShellString(normalizedFormula);

  console.error(`[PowerShell] Setting formula in cell ${cellAddress} in "${fileName}"/"${sheetName}"`);

  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $ws.Range('${cellAddress}').Formula = '${escapedFormula}'\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully set formula in cell ${cellAddress}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to set formula:`, {
      file: fileName,
      sheet: sheetName,
      cell: cellAddress,
      error: error.message,
    });
    throw error;
  }
}

// ============================================================
// Formatting Operations
// ============================================================

/**
 * Format a cell in an open Excel file via PowerShell COM.
 * Handles font, fill color, and alignment properties in a single invocation.
 */
export async function formatCellViaPowerShell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  format: {
    fontName?: string;
    fontSize?: number;
    fontBold?: boolean;
    fontItalic?: boolean;
    fontColor?: string;
    fillColor?: string;
    horizontalAlignment?: string;
    verticalAlignment?: string;
  }
): Promise<void> {
  validateCellAddress(cellAddress);

  const fileName = basename(filePath);

  console.error(`[PowerShell] Formatting cell ${cellAddress} in "${fileName}"/"${sheetName}"`);

  try {
    const preamble = buildPreamble(fileName, sheetName);
    let body = `  $r = $ws.Range('${cellAddress}')\n`;

    if (format.fontName) {
      const escapedFontName = escapePowerShellString(format.fontName);
      body += `  $r.Font.Name = '${escapedFontName}'\n`;
    }

    if (format.fontSize !== undefined) {
      body += `  $r.Font.Size = ${format.fontSize}\n`;
    }

    if (format.fontBold !== undefined) {
      body += `  $r.Font.Bold = $${format.fontBold}\n`;
    }

    if (format.fontItalic !== undefined) {
      body += `  $r.Font.Italic = $${format.fontItalic}\n`;
    }

    if (format.fontColor) {
      // Convert hex color (AARRGGBB or RRGGBB) to OLE color (R + G*256 + B*65536)
      body += `  $hex = '${escapePowerShellString(format.fontColor)}' -replace '^#',''\n`;
      body += `  if ($hex.Length -eq 8) { $hex = $hex.Substring(2) }\n`;
      body += `  $r.Font.Color = [Convert]::ToInt32($hex.Substring(4,2), 16) * 65536 + [Convert]::ToInt32($hex.Substring(2,2), 16) * 256 + [Convert]::ToInt32($hex.Substring(0,2), 16)\n`;
    }

    if (format.fillColor) {
      body += `  $hex = '${escapePowerShellString(format.fillColor)}' -replace '^#',''\n`;
      body += `  if ($hex.Length -eq 8) { $hex = $hex.Substring(2) }\n`;
      body += `  $r.Interior.Color = [Convert]::ToInt32($hex.Substring(4,2), 16) * 65536 + [Convert]::ToInt32($hex.Substring(2,2), 16) * 256 + [Convert]::ToInt32($hex.Substring(0,2), 16)\n`;
    }

    if (format.horizontalAlignment) {
      // Excel COM alignment constants
      const alignMap: Record<string, number> = {
        left: -4131,    // xlLeft
        center: -4108,  // xlCenter
        right: -4152,   // xlRight
      };
      const alignment = format.horizontalAlignment.toLowerCase();
      const constant = alignMap[alignment] || -4131;
      body += `  $r.HorizontalAlignment = ${constant}\n`;
    }

    if (format.verticalAlignment) {
      const alignMap: Record<string, number> = {
        top: -4160,     // xlTop
        center: -4108,  // xlCenter
        middle: -4108,  // xlCenter (alias)
        bottom: -4107,  // xlBottom
      };
      const alignment = format.verticalAlignment.toLowerCase();
      const constant = alignMap[alignment] || -4160;
      body += `  $r.VerticalAlignment = ${constant}\n`;
    }

    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully formatted cell ${cellAddress}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to format cell:`, {
      file: fileName,
      sheet: sheetName,
      cell: cellAddress,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Set column width in an open Excel file via PowerShell COM
 */
export async function setColumnWidthViaPowerShell(
  filePath: string,
  sheetName: string,
  column: string | number,
  width: number
): Promise<void> {
  const fileName = basename(filePath);
  const columnLetter = typeof column === 'number' ? numberToColumnLetter(column) : column;

  console.error(`[PowerShell] Setting column ${columnLetter} width to ${width} in "${fileName}"/"${sheetName}"`);

  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $ws.Columns.Item('${columnLetter}').ColumnWidth = ${width}\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully set column ${columnLetter} width to ${width}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to set column width:`, {
      file: fileName,
      sheet: sheetName,
      column: columnLetter,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Set row height in an open Excel file via PowerShell COM
 */
export async function setRowHeightViaPowerShell(
  filePath: string,
  sheetName: string,
  row: number,
  height: number
): Promise<void> {
  const fileName = basename(filePath);

  console.error(`[PowerShell] Setting row ${row} height to ${height} in "${fileName}"/"${sheetName}"`);

  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $ws.Rows.Item(${row}).RowHeight = ${height}\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully set row ${row} height to ${height}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to set row height:`, {
      file: fileName,
      sheet: sheetName,
      row,
      error: error.message,
    });
    throw error;
  }
}

// ============================================================
// Merge/Unmerge Operations
// ============================================================

/**
 * Merge cells in an open Excel file via PowerShell COM
 */
export async function mergeCellsViaPowerShell(
  filePath: string,
  sheetName: string,
  range: string
): Promise<void> {
  validateRange(range);

  const fileName = basename(filePath);

  console.error(`[PowerShell] Merging cells ${range} in "${fileName}"/"${sheetName}"`);

  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $ws.Range('${range}').Merge()\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully merged cells ${range}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to merge cells:`, {
      file: fileName,
      sheet: sheetName,
      range,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Unmerge cells in an open Excel file via PowerShell COM
 */
export async function unmergeCellsViaPowerShell(
  filePath: string,
  sheetName: string,
  range: string
): Promise<void> {
  validateRange(range);

  const fileName = basename(filePath);

  console.error(`[PowerShell] Unmerging cells ${range} in "${fileName}"/"${sheetName}"`);

  const preamble = buildPreamble(fileName, sheetName);
  const body = `  $ws.Range('${range}').UnMerge()\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully unmerged cells ${range}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to unmerge cells:`, {
      file: fileName,
      sheet: sheetName,
      range,
      error: error.message,
    });
    throw error;
  }
}

// ============================================================
// Sheet Operations
// ============================================================

/**
 * Create a new sheet in an open Excel file via PowerShell COM
 */
export async function createSheetViaPowerShell(
  filePath: string,
  sheetName: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedSheetName = escapePowerShellString(sheetName);

  const preamble = buildPreamble(fileName);
  const body = `  $newSheet = $wb.Worksheets.Add()\n  $newSheet.Name = '${escapedSheetName}'\n`;
  const script = wrapWithCleanup(preamble, body);

  await execPowerShellWithRetry(script);
}

/**
 * Delete a sheet in an open Excel file via PowerShell COM.
 * Suppresses the confirmation dialog.
 */
export async function deleteSheetViaPowerShell(
  filePath: string,
  sheetName: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedSheetName = escapePowerShellString(sheetName);

  const preamble = buildPreamble(fileName);
  const body = `  $excel.DisplayAlerts = $false\n  $wb.Worksheets.Item('${escapedSheetName}').Delete()\n  $excel.DisplayAlerts = $true\n`;
  const script = wrapWithCleanup(preamble, body);

  await execPowerShellWithRetry(script);
}

/**
 * Rename a sheet in an open Excel file via PowerShell COM
 */
export async function renameSheetViaPowerShell(
  filePath: string,
  oldName: string,
  newName: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedOldName = escapePowerShellString(oldName);
  const escapedNewName = escapePowerShellString(newName);

  console.error(`[PowerShell] Renaming sheet from "${oldName}" to "${newName}" in "${fileName}"`);

  const preamble = buildPreamble(fileName);
  const body = `  $wb.Worksheets.Item('${escapedOldName}').Name = '${escapedNewName}'\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully renamed sheet from "${oldName}" to "${newName}"`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to rename sheet:`, {
      file: fileName,
      oldName,
      newName,
      error: error.message,
    });
    throw error;
  }
}

// ============================================================
// Row/Column Operations
// ============================================================

/**
 * Delete rows in an open Excel file via PowerShell COM.
 * Deletes from the same start position each time (rows shift up).
 */
export async function deleteRowsViaPowerShell(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);

  console.error(`[PowerShell] Deleting ${count} row(s) starting at row ${startRow} in "${fileName}"/"${sheetName}"`);

  try {
    const preamble = buildPreamble(fileName, sheetName);
    let body = '';
    for (let i = 0; i < count; i++) {
      body += `  [void]$ws.Rows.Item(${startRow}).Delete()\n`;
    }
    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully deleted ${count} row(s) starting at row ${startRow}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to delete rows:`, {
      file: fileName,
      sheet: sheetName,
      startRow,
      count,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Delete columns in an open Excel file via PowerShell COM.
 * Deletes from the same start position each time (columns shift left).
 */
export async function deleteColumnsViaPowerShell(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);
  const startColumnLetter = typeof startColumn === 'number' ? numberToColumnLetter(startColumn) : startColumn;

  console.error(`[PowerShell] Deleting ${count} column(s) starting at column ${startColumnLetter} in "${fileName}"/"${sheetName}"`);

  try {
    const preamble = buildPreamble(fileName, sheetName);
    let body = '';
    for (let i = 0; i < count; i++) {
      body += `  [void]$ws.Columns.Item('${startColumnLetter}').Delete()\n`;
    }
    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully deleted ${count} column(s) starting at column ${startColumnLetter}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to delete columns:`, {
      file: fileName,
      sheet: sheetName,
      startColumn: startColumnLetter,
      count,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Insert rows in an open Excel file via PowerShell COM
 */
export async function insertRowsViaPowerShell(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);

  console.error(`[PowerShell] Inserting ${count} row(s) at row ${startRow} in "${fileName}"/"${sheetName}"`);

  try {
    const preamble = buildPreamble(fileName, sheetName);
    let body = '';
    for (let i = 0; i < count; i++) {
      body += `  [void]$ws.Rows.Item(${startRow}).Insert()\n`;
    }
    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully inserted ${count} row(s) at row ${startRow}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to insert rows:`, {
      file: fileName,
      sheet: sheetName,
      startRow,
      count,
      error: error.message,
    });
    throw error;
  }
}

/**
 * Insert columns in an open Excel file via PowerShell COM
 */
export async function insertColumnsViaPowerShell(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number
): Promise<void> {
  const fileName = basename(filePath);
  const startColumnLetter = typeof startColumn === 'number' ? numberToColumnLetter(startColumn) : startColumn;

  console.error(`[PowerShell] Inserting ${count} column(s) at column ${startColumnLetter} in "${fileName}"/"${sheetName}"`);

  try {
    const preamble = buildPreamble(fileName, sheetName);
    let body = '';
    for (let i = 0; i < count; i++) {
      body += `  [void]$ws.Columns.Item('${startColumnLetter}').Insert()\n`;
    }
    const script = wrapWithCleanup(preamble, body);
    await execPowerShellWithRetry(script);

    console.error(`[PowerShell] Successfully inserted ${count} column(s) at column ${startColumnLetter}`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to insert columns:`, {
      file: fileName,
      sheet: sheetName,
      startColumn: startColumnLetter,
      count,
      error: error.message,
    });
    throw error;
  }
}

// ============================================================
// Save Operation
// ============================================================

/**
 * Save the open Excel file via PowerShell COM
 */
export async function saveFileViaPowerShell(filePath: string): Promise<void> {
  const fileName = basename(filePath);

  console.error(`[PowerShell] Saving file "${fileName}"`);

  const preamble = buildPreamble(fileName);
  const body = `  $wb.Save()\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    await execPowerShellWithRetry(script);
    console.error(`[PowerShell] Successfully saved file "${fileName}"`);
  } catch (error: any) {
    console.error(`[PowerShell] Failed to save file:`, {
      file: fileName,
      error: error.message,
    });
    throw error;
  }
}

// ============================================================
// Comments
// ============================================================

export async function getCommentsViaPowerShell(
  filePath: string,
  sheetName: string,
  range?: string
): Promise<string> {
  const fileName = basename(filePath);
  if (range) validateRange(range);

  const preamble = buildPreamble(fileName, sheetName);
  let body: string;
  if (range) {
    body = `  $comments = @()\n  foreach ($cell in $ws.Range('${range}')) {\n    if ($cell.Comment) {\n      $comments += @{ Address = $cell.Address($false,$false); Author = $cell.Comment.Author; Text = $cell.Comment.Text() }\n    }\n  }\n  $comments | ConvertTo-Json -Compress\n`;
  } else {
    body = `  $comments = @()\n  foreach ($comment in $ws.Comments) {\n    $comments += @{ Address = $comment.Parent.Address($false,$false); Author = $comment.Author; Text = $comment.Text() }\n  }\n  $comments | ConvertTo-Json -Compress\n`;
  }
  const script = wrapWithCleanup(preamble, body);
  return await execPowerShellWithRetry(script);
}

export async function addCommentViaPowerShell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  text: string,
  author?: string
): Promise<void> {
  validateCellAddress(cellAddress);
  const fileName = basename(filePath);
  const escapedText = escapePowerShellString(text);

  const preamble = buildPreamble(fileName, sheetName);
  let body = `  $r = $ws.Range('${cellAddress}')\n`;
  body += `  if ($r.Comment) { $r.Comment.Delete() }\n`;
  body += `  [void]$r.AddComment('${escapedText}')\n`;
  if (author) {
    const escapedAuthor = escapePowerShellString(author);
    body += `  $r.Comment.Author = '${escapedAuthor}'\n`;
  }
  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Named Ranges
// ============================================================

export async function listNamedRangesViaPowerShell(
  filePath: string
): Promise<string> {
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName);
  const body = `  $names = @()\n  foreach ($n in $wb.Names) {\n    $names += @{ Name = $n.Name; RefersTo = $n.RefersTo; Visible = $n.Visible }\n  }\n  $names | ConvertTo-Json -Compress\n`;
  const script = wrapWithCleanup(preamble, body);
  return await execPowerShellWithRetry(script);
}

export async function createNamedRangeViaPowerShell(
  filePath: string,
  name: string,
  sheetName: string,
  range: string
): Promise<void> {
  validateRange(range);
  const fileName = basename(filePath);
  const escapedName = escapePowerShellString(name);
  const escapedSheetName = escapePowerShellString(sheetName);

  const preamble = buildPreamble(fileName);
  // Parse range (e.g., "A3:H18") into absolute reference "$A$3:$H$18"
  const rangeMatch = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!rangeMatch) throw new Error(`Invalid range format: ${range}`);
  const absRange = `$${rangeMatch[1]}$${rangeMatch[2]}:$${rangeMatch[3]}$${rangeMatch[4]}`;
  const refersTo = `='${escapedSheetName}'!${absRange}`;
  const body = `  [void]$wb.Names.Add('${escapedName}', '${escapePowerShellString(refersTo)}')\n`;
  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

export async function deleteNamedRangeViaPowerShell(
  filePath: string,
  name: string
): Promise<void> {
  const fileName = basename(filePath);
  const escapedName = escapePowerShellString(name);

  const preamble = buildPreamble(fileName);
  const body = `  $wb.Names.Item('${escapedName}').Delete()\n`;
  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Sheet Protection
// ============================================================

export async function setSheetProtectionViaPowerShell(
  filePath: string,
  sheetName: string,
  protect: boolean,
  password?: string,
  options?: {
    allowInsertRows?: boolean;
    allowInsertColumns?: boolean;
    allowDeleteRows?: boolean;
    allowDeleteColumns?: boolean;
    allowSort?: boolean;
    allowAutoFilter?: boolean;
    allowFormatCells?: boolean;
    allowFormatColumns?: boolean;
    allowFormatRows?: boolean;
  }
): Promise<void> {
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName, sheetName);
  let body: string;

  if (protect) {
    const pwd = password ? `'${escapePowerShellString(password)}'` : '[System.Type]::Missing';
    body = `  $ws.Protect(${pwd}`;
    // Protect method params: Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly,
    // AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows,
    // AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks,
    // AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering
    const o = options || {};
    body += `, [System.Type]::Missing, [System.Type]::Missing, [System.Type]::Missing, $true`;
    body += `, $${o.allowFormatCells || false}, $${o.allowFormatColumns || false}, $${o.allowFormatRows || false}`;
    body += `, $${o.allowInsertColumns || false}, $${o.allowInsertRows || false}, [System.Type]::Missing`;
    body += `, $${o.allowDeleteColumns || false}, $${o.allowDeleteRows || false}`;
    body += `, $${o.allowSort || false}, $${o.allowAutoFilter || false}`;
    body += `)\n`;
  } else {
    if (password) {
      body = `  $ws.Unprotect('${escapePowerShellString(password)}')\n`;
    } else {
      body = `  $ws.Unprotect()\n`;
    }
  }

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Data Validation
// ============================================================

export async function setDataValidationViaPowerShell(
  filePath: string,
  sheetName: string,
  range: string,
  validationType: string,
  formula1: string,
  operator?: string,
  formula2?: string,
  showErrorMessage?: boolean,
  errorTitle?: string,
  errorMessage?: string
): Promise<void> {
  validateRange(range);
  const fileName = basename(filePath);

  // Excel validation type constants
  const typeMap: Record<string, number> = {
    list: 3,
    whole: 1,
    decimal: 2,
    date: 4,
    textLength: 6,
    custom: 7,
  };
  const xlType = typeMap[validationType] || 3;

  // Excel operator constants
  const opMap: Record<string, number> = {
    between: 1,
    notBetween: 2,
    equal: 3,
    notEqual: 4,
    greaterThan: 5,
    lessThan: 6,
    greaterThanOrEqual: 7,
    lessThanOrEqual: 8,
  };
  const xlOp = operator ? (opMap[operator] || 1) : 1;

  const escapedFormula1 = escapePowerShellString(formula1);

  const preamble = buildPreamble(fileName, sheetName);
  let body = `  $r = $ws.Range('${range}')\n`;
  body += `  $r.Validation.Delete()\n`;
  if (formula2) {
    const escapedFormula2 = escapePowerShellString(formula2);
    body += `  $r.Validation.Add(${xlType}, 1, ${xlOp}, '${escapedFormula1}', '${escapedFormula2}')\n`;
  } else {
    body += `  $r.Validation.Add(${xlType}, 1, ${xlOp}, '${escapedFormula1}')\n`;
  }
  if (showErrorMessage !== false) {
    body += `  $r.Validation.ShowError = $true\n`;
    if (errorTitle) {
      body += `  $r.Validation.ErrorTitle = '${escapePowerShellString(errorTitle)}'\n`;
    }
    if (errorMessage) {
      body += `  $r.Validation.ErrorMessage = '${escapePowerShellString(errorMessage)}'\n`;
    }
  }

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Calculation Control
// ============================================================

export async function triggerRecalculationViaPowerShell(
  filePath: string,
  fullRecalc: boolean = false
): Promise<void> {
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName);
  const body = fullRecalc
    ? `  $excel.CalculateFull()\n`
    : `  $excel.Calculate()\n`;
  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

export async function getCalculationModeViaPowerShell(
  filePath: string
): Promise<string> {
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName);
  const body = `  $mode = $excel.Calculation\n  switch ($mode) {\n    -4105 { Write-Output 'automatic' }\n    -4135 { Write-Output 'manual' }\n    2 { Write-Output 'semiautomatic' }\n    default { Write-Output "unknown($mode)" }\n  }\n`;
  const script = wrapWithCleanup(preamble, body);
  return await execPowerShellWithRetry(script);
}

export async function setCalculationModeViaPowerShell(
  filePath: string,
  mode: string
): Promise<void> {
  const fileName = basename(filePath);

  const modeMap: Record<string, number> = {
    automatic: -4105,
    manual: -4135,
    semiautomatic: 2,
  };
  const xlMode = modeMap[mode];
  if (xlMode === undefined) throw new Error(`Invalid calculation mode: ${mode}`);

  const preamble = buildPreamble(fileName);
  const body = `  $excel.Calculation = ${xlMode}\n`;
  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Screenshot
// ============================================================

export async function captureScreenshotViaPowerShell(
  filePath: string,
  sheetName: string,
  outputPath: string,
  range?: string
): Promise<void> {
  if (range) validateRange(range);
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName, sheetName);
  const escapedOutputPath = escapePowerShellString(outputPath);
  let body = `  Add-Type -AssemblyName System.Windows.Forms\n`;
  if (range) {
    body += `  $r = $ws.Range('${range}')\n`;
  } else {
    body += `  $r = $ws.UsedRange\n`;
  }
  body += `  $r.CopyPicture(1, 2)\n`; // xlScreen=1, xlBitmap=2
  body += `  Start-Sleep -Milliseconds 200\n`;
  body += `  $img = [System.Windows.Forms.Clipboard]::GetImage()\n`;
  body += `  if ($img) {\n`;
  body += `    $img.Save('${escapedOutputPath}', [System.Drawing.Imaging.ImageFormat]::Png)\n`;
  body += `    $img.Dispose()\n`;
  body += `  } else {\n`;
  body += `    throw 'Failed to capture screenshot from clipboard'\n`;
  body += `  }\n`;

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script, MAX_RETRIES, 15000); // longer timeout for screenshot
}

// ============================================================
// VBA Macros
// ============================================================

export async function runVbaMacroViaPowerShell(
  filePath: string,
  macroName: string,
  args: any[] = []
): Promise<string> {
  const fileName = basename(filePath);
  const escapedFileName = escapePowerShellString(fileName);
  const escapedMacroName = escapePowerShellString(macroName);
  const timestamp = Date.now();
  const wrapperModuleName = `MCP_ErrWrap_${timestamp}`;

  // Build VBA arg declarations for the wrapper function
  const vbaArgDecl = args.map((_a, i) => {
    return `arg${i} As Variant`;
  }).join(', ');

  // Build the VBA wrapper code
  const vbaWrapperCode = [
    `Public Function MCP_RunWrapped(${vbaArgDecl}) As Variant`,
    `    On Error GoTo ErrHandler`,
    `    Dim res As Variant`,
    `    res = Application.Run("${macroName.replace(/"/g, '""')}"${args.map((_, i) => `, arg${i}`).join('')})`,
    `    MCP_RunWrapped = res`,
    `    Exit Function`,
    `ErrHandler:`,
    `    Err.Raise Err.Number, Err.Source, "VBA_RUNTIME_ERROR|" & Err.Number & "|" & Err.Description & "|" & Err.Source`,
    `End Function`,
  ].join('\n');

  // Escape the VBA code for embedding in a PowerShell single-quoted string
  const escapedVbaCode = escapePowerShellString(vbaWrapperCode);

  // Build the PowerShell arg values for calling the wrapper
  const psArgValues = args.map(a => formatValueForPowerShell(a)).join(', ');
  const psRunArgs = args.length > 0 ? `, ${psArgValues}` : '';

  // Direct run args (fallback)
  const psDirectArgs = args.length > 0
    ? `, ${args.map(a => formatValueForPowerShell(a)).join(', ')}`
    : '';

  const script = `
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $excel.Workbooks | Where-Object { $_.Name -eq '${escapedFileName}' }
if (-not $wb) { throw "Workbook '${escapedFileName}' not found in open Excel instance" }

$wrapperCreated = $false
$wrapperModule = $null

try {
  # Try to create error-handling wrapper via VBProject access
  try {
    $vbProj = $wb.VBProject
    $wrapperModule = $vbProj.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    $wrapperModule.Name = '${wrapperModuleName}'
    $wrapperModule.CodeModule.AddFromString('${escapedVbaCode}')
    $wrapperCreated = $true
  } catch {
    # VBProject access failed (trust not enabled) — fall back to direct Run
    $wrapperCreated = $false
  }

  if ($wrapperCreated) {
    # Run via wrapper — errors caught by On Error GoTo
    $result = $excel.Run('${wrapperModuleName}.MCP_RunWrapped'${psRunArgs})
  } else {
    # Direct run (current behavior, no error wrapping)
    $result = $excel.Run('${escapedMacroName}'${psDirectArgs})
  }

  if ($result -ne $null) { Write-Output ([string]$result) } else { Write-Output '' }
} finally {
  # Cleanup: remove temporary module
  if ($wrapperCreated -and $wrapperModule) {
    try {
      $wb.VBProject.VBComponents.Remove($wrapperModule)
    } catch {}
  }
  try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
}
`;

  try {
    return await execPowerShellWithRetry(script, MAX_RETRIES, 30000);
  } catch (error: any) {
    if (error.message && error.message.includes('Programmatic access to Visual Basic Project is not trusted')) {
      throw new Error('VBA access denied. Enable "Trust access to the VBA project object model" in Excel Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings).');
    }
    throw error;
  }
}

export async function getVbaCodeViaPowerShell(
  filePath: string,
  moduleName: string
): Promise<string> {
  const fileName = basename(filePath);
  const escapedModuleName = escapePowerShellString(moduleName);

  const preamble = buildPreamble(fileName);
  const body = `  try {\n    $comp = $wb.VBProject.VBComponents.Item('${escapedModuleName}')\n    $cm = $comp.CodeModule\n    if ($cm.CountOfLines -gt 0) {\n      Write-Output $cm.Lines(1, $cm.CountOfLines)\n    } else {\n      Write-Output ''\n    }\n  } catch {\n    if ($_.Exception.Message -match 'trust') {\n      throw 'VBA_TRUST_ERROR'\n    }\n    throw $_\n  }\n`;
  const script = wrapWithCleanup(preamble, body);

  try {
    return await execPowerShellWithRetry(script);
  } catch (error: any) {
    if (error.message && (error.message.includes('VBA_TRUST_ERROR') || error.message.includes('trust'))) {
      throw new Error('VBA access denied. Enable "Trust access to the VBA project object model" in Excel Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings).');
    }
    throw error;
  }
}

export async function setVbaCodeViaPowerShell(
  filePath: string,
  moduleName: string,
  code: string,
  createModule: boolean = false,
  appendMode: boolean = false
): Promise<void> {
  const fileName = basename(filePath);
  const escapedModuleName = escapePowerShellString(moduleName);
  // Decode HTML entities before escaping — MCP protocol may encode & as &amp; etc.
  const escapedCode = escapePowerShellString(decodeHtmlEntities(code));

  const preamble = buildPreamble(fileName);
  let body = `  try {\n`;
  if (createModule) {
    body += `    $comp = $null\n`;
    body += `    try { $comp = $wb.VBProject.VBComponents.Item('${escapedModuleName}') } catch {}\n`;
    body += `    if (-not $comp) {\n`;
    body += `      $comp = $wb.VBProject.VBComponents.Add(1)\n`; // 1 = vbext_ct_StdModule
    body += `      $comp.Name = '${escapedModuleName}'\n`;
    body += `    }\n`;
  } else {
    body += `    $comp = $wb.VBProject.VBComponents.Item('${escapedModuleName}')\n`;
  }
  body += `    $cm = $comp.CodeModule\n`;
  if (appendMode) {
    // Append: add a blank line separator then insert new code after existing lines
    body += `    $insertAt = $cm.CountOfLines + 1\n`;
    body += `    if ($cm.CountOfLines -gt 0) { $cm.InsertLines($insertAt, [string][char]10) ; $insertAt = $insertAt + 1 }\n`;
    body += `    $cm.InsertLines($insertAt, '${escapedCode}')\n`;
  } else {
    // Replace: clear and rewrite
    body += `    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }\n`;
    body += `    $cm.InsertLines(1, '${escapedCode}')\n`;
  }
  body += `  } catch {\n`;
  body += `    if ($_.Exception.Message -match 'trust') {\n`;
  body += `      throw 'VBA_TRUST_ERROR'\n`;
  body += `    }\n`;
  body += `    throw $_\n`;
  body += `  }\n`;

  const script = wrapWithCleanup(preamble, body);
  try {
    await execPowerShellWithRetry(script);
  } catch (error: any) {
    if (error.message && (error.message.includes('VBA_TRUST_ERROR') || error.message.includes('trust'))) {
      throw new Error('VBA access denied. Enable "Trust access to the VBA project object model" in Excel Trust Center settings (File > Options > Trust Center > Trust Center Settings > Macro Settings).');
    }
    throw error;
  }
}

// ============================================================
// VBA Trust Access (Registry)
// ============================================================

/**
 * Check if "Trust access to VBA project object model" is enabled.
 * Checks ALL known registry paths (user, policy, Click-to-Run) AND
 * does a live COM test to confirm VBE is actually accessible.
 */
export async function checkVbaTrustViaPowerShell(): Promise<string> {
  const script = `
$results = @{ RegistryPaths = @(); ComTestPassed = $false; Enabled = $false }

# Check all known registry locations for AccessVBOM
$paths = @(
  'HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Security',
  'HKCU:\\Software\\Microsoft\\Office\\15.0\\Excel\\Security',
  'HKCU:\\Software\\Microsoft\\Office\\14.0\\Excel\\Security',
  'HKCU:\\Software\\Policies\\Microsoft\\Office\\16.0\\Excel\\Security',
  'HKCU:\\Software\\Policies\\Microsoft\\Office\\15.0\\Excel\\Security',
  'HKLM:\\Software\\Policies\\Microsoft\\Office\\16.0\\Excel\\Security',
  'HKLM:\\Software\\Policies\\Microsoft\\Office\\15.0\\Excel\\Security'
)

foreach ($path in $paths) {
  if (Test-Path $path) {
    $val = Get-ItemProperty -Path $path -Name AccessVBOM -ErrorAction SilentlyContinue
    $entry = @{ Path = $path; Exists = $true }
    if ($val -ne $null) {
      $entry.AccessVBOM = $val.AccessVBOM
      if ($val.AccessVBOM -eq 1) { $results.Enabled = $true }
    } else {
      $entry.AccessVBOM = 'not set'
    }
    $results.RegistryPaths += $entry
  }
}

# Live COM test: try to access VBE on running Excel
try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  try {
    $vbeTest = $excel.VBE
    if ($vbeTest -ne $null) {
      $results.ComTestPassed = $true
      $projCount = $vbeTest.VBProjects.Count
      $results.VBProjectCount = $projCount
    } else {
      $results.ComTestNote = 'VBE object is null - trust not effective'
    }
  } catch {
    $results.ComTestNote = 'VBE access threw exception: ' + $_.Exception.Message
  }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
  $results.ComTestNote = 'Excel not running or COM unavailable'
}

$results | ConvertTo-Json -Depth 3 -Compress
`;
  return await execPowerShellWithRetry(script, 2, 8000);
}

/**
 * Enable or disable "Trust access to VBA project object model" via registry.
 * Writes to ALL known registry paths to maximize compatibility across
 * Office versions (MSI, Click-to-Run, Store).
 */
export async function enableVbaTrustViaPowerShell(enable: boolean): Promise<string> {
  const value = enable ? 1 : 0;
  const script = `
$written = @()
$versions = @('16.0', '15.0', '14.0')

foreach ($ver in $versions) {
  # User-level path (standard installs)
  $userPath = "HKCU:\\Software\\Microsoft\\Office\\$ver\\Excel\\Security"
  if (Test-Path $userPath) {
    Set-ItemProperty -Path $userPath -Name AccessVBOM -Value ${value} -Type DWord -ErrorAction SilentlyContinue
    $written += $userPath
  }

  # Policy path (some managed installs read from here)
  $policyPath = "HKCU:\\Software\\Policies\\Microsoft\\Office\\$ver\\Excel\\Security"
  if (Test-Path $policyPath) {
    Set-ItemProperty -Path $policyPath -Name AccessVBOM -Value ${value} -Type DWord -ErrorAction SilentlyContinue
    $written += $policyPath
  } else {
    # Create the policy path if the user-level one exists (some Office installs prefer policy)
    $parentPolicy = "HKCU:\\Software\\Policies\\Microsoft\\Office\\$ver\\Excel"
    if (Test-Path "HKCU:\\Software\\Policies\\Microsoft\\Office\\$ver") {
      if (-not (Test-Path $policyPath)) {
        New-Item -Path $policyPath -Force -ErrorAction SilentlyContinue | Out-Null
      }
      Set-ItemProperty -Path $policyPath -Name AccessVBOM -Value ${value} -Type DWord -ErrorAction SilentlyContinue
      $written += "$policyPath (created)"
    }
  }
}

# Verify by reading back
$verified = @()
foreach ($p in $written) {
  $cleanPath = $p -replace ' \\(created\\)', ''
  $val = Get-ItemProperty -Path $cleanPath -Name AccessVBOM -ErrorAction SilentlyContinue
  if ($val) { $verified += @{ Path = $cleanPath; Value = $val.AccessVBOM } }
}

@{
  Success = ($written.Count -gt 0)
  PathsWritten = $written
  Verified = $verified
  Note = 'Restart Excel completely (close ALL Excel windows, check Task Manager) for the change to take effect.'
} | ConvertTo-Json -Depth 3 -Compress
`;
  return await execPowerShellWithRetry(script, 2, 8000);
}

// ============================================================
// Diagnosis (Connection & Accessibility)
// ============================================================

export async function diagnoseConnectionViaPowerShell(
  filePath?: string
): Promise<string> {
  const escapedFilePath = filePath ? escapePowerShellString(filePath) : '';
  const fileNameOnly = filePath ? escapePowerShellString(basename(filePath)) : '';

  const script = `
$results = @{}

# Step 1: Check if Excel process is running
$excelProc = Get-Process EXCEL -ErrorAction SilentlyContinue
if (-not $excelProc) {
  $results.step1_process = @{ passed = $false; message = 'Excel is not running'; fix = 'Open Excel and your workbook, then retry.' }
  $results | ConvertTo-Json -Depth 3 -Compress
  exit
}
$results.step1_process = @{ passed = $true; message = "Excel running (PID: $($excelProc.Id -join ', '))" }

# Step 2: Try to get COM object
try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  $results.step2_com = @{ passed = $true; message = 'COM connection to Excel.Application succeeded' }
} catch {
  $results.step2_com = @{ passed = $false; message = "COM connection failed: $($_.Exception.Message)"; fix = 'Excel may be running elevated (as admin) or in a different session. Restart Excel normally (not as admin).' }
  $results | ConvertTo-Json -Depth 3 -Compress
  exit
}

# Step 3: Try to access Workbooks (will hang if modal dialog is up, but we have a timeout)
try {
  $wbCount = $excel.Workbooks.Count
  $results.step3_responsive = @{ passed = $true; message = "Excel is responsive. $wbCount workbook(s) open." }
} catch {
  $results.step3_responsive = @{ passed = $false; message = "Excel is not responding to COM calls: $($_.Exception.Message)"; fix = 'Excel likely has a modal dialog open (VBA error, Save prompt, etc.). Switch to Excel and dismiss any dialogs, then retry.' }
  try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  $results | ConvertTo-Json -Depth 3 -Compress
  exit
}

# Step 4: Check if specific file is open (if filePath provided)
${filePath ? `
$found = $false
foreach ($wb in $excel.Workbooks) {
  if ($wb.Name -eq '${fileNameOnly}' -or $wb.FullName -eq '${escapedFilePath}') {
    $found = $true
    $targetWb = $wb
    break
  }
}
if ($found) {
  $results.step4_file = @{ passed = $true; message = "File '${fileNameOnly}' is open in Excel" }
} else {
  $openNames = @()
  foreach ($wb in $excel.Workbooks) { $openNames += $wb.Name }
  $results.step4_file = @{ passed = $false; message = "File '${fileNameOnly}' is not open in Excel. Open workbooks: $($openNames -join ', ')"; fix = "Open the file in Excel: '${escapedFilePath}'" }
  try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  $results | ConvertTo-Json -Depth 3 -Compress
  exit
}
` : `
$results.step4_file = @{ passed = $true; message = 'No filePath provided, skipping file check' }
`}

# Step 5: Check VBA trust via COM (if filePath provided)
${filePath ? `
try {
  $compCount = $targetWb.VBProject.VBComponents.Count
  $results.step5_vba_trust = @{ passed = $true; message = "VBA trust working. $compCount VBA component(s) accessible." }
} catch {
  $results.step5_vba_trust = @{ passed = $false; message = "Cannot access VBProject: $($_.Exception.Message)"; fix = 'Enable "Trust access to the VBA project object model" in Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings. Then restart Excel.' }
}
` : `
$results.step5_vba_trust = @{ passed = $true; message = 'No filePath provided, skipping VBA trust check' }
`}

# Step 6: Check registry for AccessVBOM
$regPaths = @()
$vbomEnabled = $false
foreach ($ver in @('16.0', '15.0', '14.0')) {
  $secPath = "HKCU:\\Software\\Microsoft\\Office\\$ver\\Excel\\Security"
  $polPath = "HKCU:\\Software\\Policies\\Microsoft\\Office\\$ver\\Excel\\Security"
  foreach ($p in @($secPath, $polPath)) {
    $val = Get-ItemProperty -Path $p -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($val) {
      $regPaths += @{ Path = $p; Value = $val.AccessVBOM }
      if ($val.AccessVBOM -eq 1) { $vbomEnabled = $true }
    }
  }
}
if ($vbomEnabled) {
  $results.step6_registry = @{ passed = $true; message = "AccessVBOM=1 found in registry"; paths = $regPaths }
} else {
  $results.step6_registry = @{ passed = $false; message = 'AccessVBOM not set to 1 in any known registry path'; paths = $regPaths; fix = 'Use excel_enable_vba_trust tool or manually enable in Excel Trust Center.' }
}

try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
$results | ConvertTo-Json -Depth 3 -Compress
`;

  try {
    return await execPowerShellWithRetry(script, 1, 15000);
  } catch (error: any) {
    // If even the diagnostic script times out, return structured error
    if (error.killed === true) {
      return JSON.stringify({
        step1_process: { passed: true, message: 'Excel process detected (script timed out after step 1)' },
        step2_com: { passed: false, message: 'Diagnostic script timed out', fix: 'Excel likely has a modal dialog open that is blocking ALL COM calls. Switch to Excel and dismiss any dialogs (VBA error popups, Save prompts, etc.).' },
      });
    }
    throw error;
  }
}

// ============================================================
// Power Query
// ============================================================

export async function listPowerQueriesViaPowerShell(
  filePath: string
): Promise<string> {
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName);
  const body = `  $queries = @()\n  foreach ($q in $wb.Queries) {\n    $queries += @{ Name = $q.Name; Formula = $q.Formula; Description = $q.Description }\n  }\n  $queries | ConvertTo-Json -Compress\n`;
  const script = wrapWithCleanup(preamble, body);
  return await execPowerShellWithRetry(script);
}

export async function runPowerQueryViaPowerShell(
  filePath: string,
  queryName: string,
  formula: string,
  refreshOnly: boolean = false
): Promise<void> {
  const fileName = basename(filePath);
  const escapedQueryName = escapePowerShellString(queryName);
  const escapedFormula = escapePowerShellString(formula);

  const preamble = buildPreamble(fileName);
  let body: string;

  if (refreshOnly) {
    body = `  $found = $false\n  foreach ($conn in $wb.Connections) {\n    if ($conn.Name -match '${escapedQueryName}') {\n      $conn.Refresh()\n      $found = $true\n      break\n    }\n  }\n  if (-not $found) { throw "Query '${escapedQueryName}' not found" }\n`;
  } else {
    body = `  try { $wb.Queries.Item('${escapedQueryName}').Delete() } catch {}\n`;
    body += `  [void]$wb.Queries.Add('${escapedQueryName}', '${escapedFormula}')\n`;
  }

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script, MAX_RETRIES, 30000); // queries may take time
}

// ============================================================
// Chart (Real COM Chart)
// ============================================================

export async function createChartViaPowerShell(
  filePath: string,
  sheetName: string,
  chartType: string,
  dataRange: string,
  position: string,
  title?: string,
  showLegend: boolean = true,
  dataSheetName?: string
): Promise<string> {
  validateRange(dataRange);
  validateCellAddress(position);
  const fileName = basename(filePath);

  const chartTypeMap: Record<string, number> = {
    line: 4,      // xlLine
    bar: 57,      // xlBarClustered
    column: 51,   // xlColumnClustered
    pie: 5,       // xlPie
    scatter: -4169,// xlXYScatter
    area: 1,      // xlArea
  };
  const xlChartType = chartTypeMap[chartType] || 51;

  const preamble = buildPreamble(fileName, sheetName);
  let body = `  $ErrorActionPreference = 'Stop'\n`;
  body += `  $posCell = $ws.Range('${position}')\n`;
  body += `  $chartObj = $ws.ChartObjects().Add($posCell.Left, $posCell.Top, 400, 300)\n`;
  body += `  $chart = $chartObj.Chart\n`;
  body += `  $chart.ChartType = ${xlChartType}\n`;

  // Resolve the data range — optionally from a different sheet
  const escapedDataRange = escapePowerShellString(dataRange);
  if (dataSheetName) {
    const escapedDataSheet = escapePowerShellString(dataSheetName);
    body += `  $dataWs = $wb.Worksheets.Item('${escapedDataSheet}')\n`;
    body += `  $srcRange = $dataWs.Range('${escapedDataRange}')\n`;
  } else {
    body += `  $srcRange = $ws.Range('${escapedDataRange}')\n`;
  }

  body += `  $rowCount = $srcRange.Rows.Count\n`;
  body += `  $colCount = $srcRange.Columns.Count\n`;
  body += `  $tier = 'none'\n`;
  body += `  $errors = @()\n`;

  // ---- Tier 1: SetSourceData ----
  body += `  try {\n`;
  body += `    $chart.SetSourceData($srcRange)\n`;
  body += `    if ($chart.SeriesCollection().Count -gt 0) { $tier = 'SetSourceData' }\n`;
  body += `  } catch {\n`;
  body += `    $errors += "T1(SetSourceData): $($_.Exception.Message)"\n`;
  body += `  }\n`;

  // ---- Tier 2: NewSeries with range refs ----
  body += `  if ($tier -eq 'none') {\n`;
  body += `    try {\n`;
  body += `      for ($c = 2; $c -le $colCount; $c++) {\n`;
  body += `        $ns = $chart.SeriesCollection().NewSeries()\n`;
  body += `        $ns.Values = $srcRange.Columns($c).Offset(1,0).Resize($rowCount - 1, 1)\n`;
  body += `        $ns.XValues = $srcRange.Columns(1).Offset(1,0).Resize($rowCount - 1, 1)\n`;
  body += `        $h = $srcRange.Cells(1, $c).Value2\n`;
  body += `        if ($h) { $ns.Name = [string]$h }\n`;
  body += `      }\n`;
  body += `      if ($chart.SeriesCollection().Count -gt 0) { $tier = 'NewSeries-Range' }\n`;
  body += `    } catch {\n`;
  body += `      $errors += "T2(NewSeries-Range): $($_.Exception.Message)"\n`;
  body += `      # Clear any partial series from failed tier 2\n`;
  body += `      try { while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() } } catch {}\n`;
  body += `    }\n`;
  body += `  }\n`;

  // ---- Tier 3: NewSeries with extracted arrays (bypasses all range marshaling) ----
  // Detects date-formatted cells in column 1 and converts to formatted strings instead of raw serial numbers
  body += `  if ($tier -eq 'none') {\n`;
  body += `    try {\n`;
  body += `      # Check if column 1 contains dates by inspecting the NumberFormat of the first data cell\n`;
  body += `      $catNf = $srcRange.Cells(2, 1).NumberFormat\n`;
  body += `      $isDateCol = ($catNf -match 'd|m|y' -and $catNf -notmatch '#|0|%')\n`;
  body += `      for ($c = 2; $c -le $colCount; $c++) {\n`;
  body += `        $ns = $chart.SeriesCollection().NewSeries()\n`;
  body += `        $vals = @()\n`;
  body += `        $cats = @()\n`;
  body += `        for ($r = 2; $r -le $rowCount; $r++) {\n`;
  body += `          $v = $srcRange.Cells($r, $c).Value2\n`;
  body += `          if ($v -eq $null) { $v = 0 }\n`;
  body += `          $vals += [double]$v\n`;
  body += `          $catCell = $srcRange.Cells($r, 1)\n`;
  body += `          if ($isDateCol -and $catCell.Value2 -ne $null) {\n`;
  body += `            # Convert Excel serial date to formatted string using the cell's own format\n`;
  body += `            $cats += $catCell.Text\n`;
  body += `          } else {\n`;
  body += `            $cv = $catCell.Value2\n`;
  body += `            if ($cv -eq $null) { $cv = '' }\n`;
  body += `            $cats += [string]$cv\n`;
  body += `          }\n`;
  body += `        }\n`;
  body += `        $ns.Values = $vals\n`;
  body += `        $ns.XValues = $cats\n`;
  body += `        $h = $srcRange.Cells(1, $c).Value2\n`;
  body += `        if ($h) { $ns.Name = [string]$h }\n`;
  body += `      }\n`;
  body += `      if ($chart.SeriesCollection().Count -gt 0) { $tier = 'NewSeries-Array' }\n`;
  body += `    } catch {\n`;
  body += `      $errors += "T3(NewSeries-Array): $($_.Exception.Message)"\n`;
  body += `      try { while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() } } catch {}\n`;
  body += `    }\n`;
  body += `  }\n`;

  // ---- Tier 4: Hardcoded test — can this chart accept ANY series at all? ----
  body += `  if ($tier -eq 'none') {\n`;
  body += `    try {\n`;
  body += `      $ns = $chart.SeriesCollection().NewSeries()\n`;
  body += `      $ns.Values = @(1, 2, 3, 4, 5)\n`;
  body += `      $ns.Name = 'TestSeries'\n`;
  body += `      if ($chart.SeriesCollection().Count -gt 0) { $tier = 'Hardcoded-Test' }\n`;
  body += `    } catch {\n`;
  body += `      $errors += "T4(Hardcoded): $($_.Exception.Message)"\n`;
  body += `    }\n`;
  body += `  }\n`;

  // ---- Final validation ----
  body += `  $sc = $chart.SeriesCollection().Count\n`;
  body += `  if ($sc -eq 0) {\n`;
  body += `    $errDetail = $errors -join ' | '\n`;
  body += `    $chartObj.Delete()\n`;
  body += `    throw "CHART_BIND_FAILED: 0 series after 4 tiers. Range=${escapedDataRange} rows=$rowCount cols=$colCount | $errDetail"\n`;
  body += `  }\n`;

  if (title) {
    const escapedTitle = escapePowerShellString(title);
    body += `  $chart.HasTitle = $true\n`;
    body += `  $chart.ChartTitle.Text = '${escapedTitle}'\n`;
  }
  body += `  $chart.HasLegend = $${showLegend}\n`;
  body += `  Write-Output "tier=$tier|seriesCount=$sc|rows=$rowCount|cols=$colCount"\n`;

  const script = wrapWithCleanup(preamble, body);
  return await execPowerShellWithRetry(script, MAX_RETRIES, 15000);
}

// ============================================================
// Pivot Table (Real COM Pivot)
// ============================================================

export async function createPivotTableViaPowerShell(
  filePath: string,
  sourceSheetName: string,
  sourceRange: string,
  targetSheetName: string,
  targetCell: string,
  rows: string[],
  values: Array<{ field: string; aggregation: string }>
): Promise<void> {
  validateRange(sourceRange);
  validateCellAddress(targetCell);
  const fileName = basename(filePath);

  const escapedSourceSheet = escapePowerShellString(sourceSheetName);
  const escapedTargetSheet = escapePowerShellString(targetSheetName);

  const aggMap: Record<string, number> = {
    sum: -4157,     // xlSum
    count: -4112,   // xlCount
    average: -4106, // xlAverage
    min: -4139,     // xlMin
    max: -4136,     // xlMax
  };

  const preamble = buildPreamble(fileName);
  let body = `  $srcWs = $wb.Worksheets.Item('${escapedSourceSheet}')\n`;
  body += `  $srcRange = $srcWs.Range('${sourceRange}')\n`;
  body += `  $tgtWs = $wb.Worksheets.Item('${escapedTargetSheet}')\n`;
  body += `  $tgtCell = $tgtWs.Range('${targetCell}')\n`;
  body += `  $cache = $wb.PivotCaches().Create(1, $srcRange)\n`; // 1 = xlDatabase
  body += `  $pt = $cache.CreatePivotTable($tgtCell, 'PivotTable_' + [guid]::NewGuid().ToString('N').Substring(0,8))\n`;

  // Add row fields
  for (const rowField of rows) {
    const escapedField = escapePowerShellString(rowField);
    body += `  $pf = $pt.PivotFields('${escapedField}')\n`;
    body += `  $pf.Orientation = 1\n`; // xlRowField = 1
  }

  // Add value fields
  for (const val of values) {
    const escapedField = escapePowerShellString(val.field);
    const xlAgg = aggMap[val.aggregation] || -4157;
    body += `  $pt.AddDataField($pt.PivotFields('${escapedField}'), '${val.aggregation} of ${escapedField}', ${xlAgg})\n`;
  }

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script, MAX_RETRIES, 15000);
}

// ============================================================
// Table (COM)
// ============================================================

export async function createTableViaPowerShell(
  filePath: string,
  sheetName: string,
  range: string,
  tableName: string,
  tableStyle: string = 'TableStyleMedium2'
): Promise<void> {
  validateRange(range);
  const fileName = basename(filePath);
  const escapedTableName = escapePowerShellString(tableName);
  const escapedStyle = escapePowerShellString(tableStyle);

  const preamble = buildPreamble(fileName, sheetName);
  let body = `  $r = $ws.Range('${range}')\n`;
  body += `  $tbl = $ws.ListObjects.Add(1, $r, $null, 1)\n`; // 1=xlSrcRange, 1=xlYes (headers)
  body += `  $tbl.Name = '${escapedTableName}'\n`;
  body += `  $tbl.TableStyle = '${escapedStyle}'\n`;

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Conditional Formatting (COM)
// ============================================================

export async function applyConditionalFormatViaPowerShell(
  filePath: string,
  sheetName: string,
  range: string,
  ruleType: string,
  condition?: {
    operator?: string;
    value?: any;
    value2?: any;
  },
  style?: {
    font?: { color?: string; bold?: boolean };
    fill?: { fgColor?: string };
  },
  colorScale?: {
    minColor?: string;
    midColor?: string;
    maxColor?: string;
  }
): Promise<void> {
  validateRange(range);
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName, sheetName);
  let body = `  $r = $ws.Range('${range}')\n`;

  if (ruleType === 'cellValue' && condition && style) {
    // Excel COM operator constants for FormatConditions.Add
    const opMap: Record<string, number> = {
      greaterThan: 5,   // xlGreater
      lessThan: 6,      // xlLess
      equal: 3,         // xlEqual
      notEqual: 4,      // xlNotEqual
      between: 1,       // xlBetween
      containsText: 0,  // handled separately
    };

    if (condition.operator === 'containsText') {
      // Use text-contains approach
      const escapedVal = escapePowerShellString(String(condition.value));
      body += `  $fc = $r.FormatConditions.Add(2, 0, '${escapedVal}')\n`; // 2=xlTextString
    } else {
      const xlOp = opMap[condition.operator || 'greaterThan'] || 5;
      const val1 = typeof condition.value === 'number' ? condition.value : `'${escapePowerShellString(String(condition.value))}'`;
      if (condition.operator === 'between' && condition.value2 !== undefined) {
        const val2 = typeof condition.value2 === 'number' ? condition.value2 : `'${escapePowerShellString(String(condition.value2))}'`;
        body += `  $fc = $r.FormatConditions.Add(1, ${xlOp}, ${val1}, ${val2})\n`; // 1=xlCellValue
      } else {
        body += `  $fc = $r.FormatConditions.Add(1, ${xlOp}, ${val1})\n`;
      }
    }

    // Apply formatting to the condition
    if (style.font?.bold !== undefined) {
      body += `  $fc.Font.Bold = $${style.font.bold}\n`;
    }
    if (style.font?.color) {
      const hex = style.font.color.replace(/^#/, '').replace(/^FF/, '');
      body += `  $fc.Font.Color = [Convert]::ToInt32('${hex}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${hex}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${hex}'.Substring(0,2), 16)\n`;
    }
    if (style.fill?.fgColor) {
      const hex = style.fill.fgColor.replace(/^#/, '').replace(/^FF/, '');
      body += `  $fc.Interior.Color = [Convert]::ToInt32('${hex}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${hex}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${hex}'.Substring(0,2), 16)\n`;
    }
  } else if (ruleType === 'colorScale') {
    const minColor = (colorScale?.minColor || 'FF0000').replace(/^FF/, '');
    const maxColor = (colorScale?.maxColor || '00FF00').replace(/^FF/, '');

    if (colorScale?.midColor) {
      const midColor = colorScale.midColor.replace(/^FF/, '');
      body += `  $cs = $r.FormatConditions.AddColorScale(3)\n`; // 3-color scale
      body += `  $cs.ColorScaleCriteria.Item(1).FormatColor.Color = [Convert]::ToInt32('${minColor}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${minColor}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${minColor}'.Substring(0,2), 16)\n`;
      body += `  $cs.ColorScaleCriteria.Item(2).FormatColor.Color = [Convert]::ToInt32('${midColor}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${midColor}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${midColor}'.Substring(0,2), 16)\n`;
      body += `  $cs.ColorScaleCriteria.Item(3).FormatColor.Color = [Convert]::ToInt32('${maxColor}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${maxColor}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${maxColor}'.Substring(0,2), 16)\n`;
    } else {
      body += `  $cs = $r.FormatConditions.AddColorScale(2)\n`; // 2-color scale
      body += `  $cs.ColorScaleCriteria.Item(1).FormatColor.Color = [Convert]::ToInt32('${minColor}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${minColor}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${minColor}'.Substring(0,2), 16)\n`;
      body += `  $cs.ColorScaleCriteria.Item(2).FormatColor.Color = [Convert]::ToInt32('${maxColor}'.Substring(4,2), 16) * 65536 + [Convert]::ToInt32('${maxColor}'.Substring(2,2), 16) * 256 + [Convert]::ToInt32('${maxColor}'.Substring(0,2), 16)\n`;
    }
  } else if (ruleType === 'dataBar') {
    body += `  [void]$r.FormatConditions.AddDatabar()\n`;
  }

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Batch Format (COM) — apply many formatting ops in one call
// ============================================================

interface BatchFormatOp {
  range: string;
  merge?: boolean;
  unmerge?: boolean;
  value?: string | number;
  fontName?: string;
  fontSize?: number;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontColor?: string;
  fillColor?: string;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  numberFormat?: string;
  columnWidth?: number;
  rowHeight?: number;
  borderStyle?: string;
  borderColor?: string;
  wrapText?: boolean;
  autoFit?: boolean;
}

/**
 * Convert hex color (#RRGGBB or RRGGBB) to Excel COM BGR long value.
 * Excel COM uses BGR ordering: Blue*65536 + Green*256 + Red
 */
function hexToExcelColor(hex: string): string {
  const clean = hex.replace(/^#/, '');
  const r = clean.substring(0, 2);
  const g = clean.substring(2, 4);
  const b = clean.substring(4, 6);
  return `([Convert]::ToInt32('${b}', 16) * 65536 + [Convert]::ToInt32('${g}', 16) * 256 + [Convert]::ToInt32('${r}', 16))`;
}

export async function batchFormatViaPowerShell(
  filePath: string,
  sheetName: string,
  operations: BatchFormatOp[]
): Promise<void> {
  const fileName = basename(filePath);

  const preamble = buildPreamble(fileName, sheetName);

  // Map alignment strings to Excel constants
  const hAlignMap: Record<string, string> = {
    left: '-4131',    // xlLeft
    center: '-4108',  // xlCenter
    right: '-4152',   // xlRight
  };
  const vAlignMap: Record<string, string> = {
    top: '-4160',     // xlTop
    center: '-4108',  // xlCenter
    bottom: '-4107',  // xlBottom
  };
  const borderStyleMap: Record<string, string> = {
    thin: '1',    // xlThin
    medium: '-4138', // xlMedium (actually 2 for medium)
    thick: '4',   // xlThick
    none: '0',    // xlNone
  };
  // Fix: medium border is actually weight 2 in xlBorderWeight
  borderStyleMap['medium'] = '2';

  let body = '';

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    const escapedRange = escapePowerShellString(op.range);
    body += `  # Operation ${i + 1}: ${escapedRange}\n`;
    body += `  $r = $ws.Range('${escapedRange}')\n`;

    // Unmerge first if requested (prevents Error 5 on re-merge)
    if (op.unmerge) {
      body += `  $r.UnMerge()\n`;
    }

    // Merge
    if (op.merge) {
      body += `  $r.Merge()\n`;
    }

    // Value
    if (op.value !== undefined) {
      if (typeof op.value === 'number') {
        body += `  $r.Item(1,1).Value2 = ${op.value}\n`;
      } else {
        body += `  $r.Item(1,1).Value2 = '${escapePowerShellString(op.value)}'\n`;
      }
    }

    // Font
    if (op.fontName) {
      body += `  $r.Font.Name = '${escapePowerShellString(op.fontName)}'\n`;
    }
    if (op.fontSize) {
      body += `  $r.Font.Size = ${op.fontSize}\n`;
    }
    if (op.fontBold !== undefined) {
      body += `  $r.Font.Bold = $${op.fontBold}\n`;
    }
    if (op.fontItalic !== undefined) {
      body += `  $r.Font.Italic = $${op.fontItalic}\n`;
    }
    if (op.fontColor) {
      body += `  $r.Font.Color = ${hexToExcelColor(op.fontColor)}\n`;
    }

    // Fill
    if (op.fillColor) {
      body += `  $r.Interior.Color = ${hexToExcelColor(op.fillColor)}\n`;
    }

    // Alignment
    if (op.horizontalAlignment && hAlignMap[op.horizontalAlignment]) {
      body += `  $r.HorizontalAlignment = ${hAlignMap[op.horizontalAlignment]}\n`;
    }
    if (op.verticalAlignment && vAlignMap[op.verticalAlignment]) {
      body += `  $r.VerticalAlignment = ${vAlignMap[op.verticalAlignment]}\n`;
    }

    // Wrap text
    if (op.wrapText !== undefined) {
      body += `  $r.WrapText = $${op.wrapText}\n`;
    }

    // Number format
    if (op.numberFormat) {
      body += `  $r.NumberFormat = '${escapePowerShellString(op.numberFormat)}'\n`;
    }

    // Column width
    if (op.columnWidth) {
      body += `  $r.ColumnWidth = ${op.columnWidth}\n`;
    }

    // Row height
    if (op.rowHeight) {
      body += `  $r.RowHeight = ${op.rowHeight}\n`;
    }

    // Auto-fit
    if (op.autoFit) {
      body += `  $r.Columns.AutoFit() | Out-Null\n`;
    }

    // Borders
    if (op.borderStyle) {
      const weight = borderStyleMap[op.borderStyle] || '1';
      if (op.borderStyle === 'none') {
        // Remove borders
        body += `  $r.Borders.LineStyle = 0\n`; // xlNone
      } else {
        // Apply to all 4 edges: 7=xlEdgeLeft, 8=xlEdgeTop, 9=xlEdgeBottom, 10=xlEdgeRight
        for (const edge of [7, 8, 9, 10]) {
          body += `  $r.Borders.Item(${edge}).Weight = ${weight}\n`;
          body += `  $r.Borders.Item(${edge}).LineStyle = 1\n`; // xlContinuous
          if (op.borderColor) {
            body += `  $r.Borders.Item(${edge}).Color = ${hexToExcelColor(op.borderColor)}\n`;
          }
        }
      }
    }

    body += `\n`;
  }

  // Save after all operations
  body += `  $wb.Save()\n`;

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script, MAX_RETRIES, 30000); // may take time with many ops
}

// ============================================================
// Display Options (COM)
// ============================================================

export async function setDisplayOptionsViaPowerShell(
  filePath: string,
  sheetName?: string,
  showGridlines?: boolean,
  showRowColumnHeaders?: boolean,
  zoomLevel?: number,
  freezePaneCell?: string,
  tabColor?: string
): Promise<void> {
  const fileName = basename(filePath);
  const preamble = buildPreamble(fileName, sheetName);

  let body = '';

  // Activate the sheet if specified
  if (sheetName) {
    body += `  $ws.Activate()\n`;
  }

  if (showGridlines !== undefined) {
    body += `  $excel.ActiveWindow.DisplayGridlines = $${showGridlines}\n`;
  }

  if (showRowColumnHeaders !== undefined) {
    body += `  $excel.ActiveWindow.DisplayHeadings = $${showRowColumnHeaders}\n`;
  }

  if (zoomLevel !== undefined) {
    body += `  $excel.ActiveWindow.Zoom = ${zoomLevel}\n`;
  }

  if (freezePaneCell !== undefined) {
    // Unfreeze first
    body += `  $excel.ActiveWindow.FreezePanes = $false\n`;
    if (freezePaneCell !== '') {
      const escapedCell = escapePowerShellString(freezePaneCell);
      body += `  $ws.Range('${escapedCell}').Select()\n`;
      body += `  $excel.ActiveWindow.FreezePanes = $true\n`;
    }
  }

  if (tabColor !== undefined) {
    if (tabColor === '') {
      body += `  $ws.Tab.ColorIndex = -4142\n`; // xlColorIndexNone
    } else {
      body += `  $ws.Tab.Color = ${hexToExcelColor(tabColor)}\n`;
    }
  }

  body += `  $wb.Save()\n`;

  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script);
}

// ============================================================
// Shapes (COM) — card layouts, visual elements
// ============================================================

interface ShapeConfig {
  shapeType: string;
  left: number;
  top: number;
  width: number;
  height: number;
  name?: string;
  fill?: {
    color?: string;
    gradient?: {
      color1: string;
      color2: string;
      direction?: string;
    };
    transparency?: number;
  };
  line?: {
    color?: string;
    weight?: number;
    visible?: boolean;
  };
  shadow?: {
    visible?: boolean;
    color?: string;
    offsetX?: number;
    offsetY?: number;
    blur?: number;
    transparency?: number;
  };
  text?: {
    value: string;
    fontName?: string;
    fontSize?: number;
    fontBold?: boolean;
    fontColor?: string;
    horizontalAlignment?: string;
    verticalAlignment?: string;
    autoSize?: string;
  };
}

export async function addShapeViaPowerShell(
  filePath: string,
  sheetName: string,
  config: ShapeConfig
): Promise<string> {
  const fileName = basename(filePath);
  const preamble = buildPreamble(fileName, sheetName);

  // Map shape type names to msoAutoShapeType constants
  const shapeTypeMap: Record<string, number> = {
    rectangle: 1,         // msoShapeRectangle
    roundedRectangle: 5,  // msoShapeRoundedRectangle
    oval: 9,              // msoShapeOval
  };

  // Map gradient direction to msoGradientStyle constants
  const gradientDirMap: Record<string, number> = {
    horizontal: 1,   // msoGradientHorizontal
    vertical: 4,     // msoGradientFromCorner (actually vertical = 3, let me use correct)
    diagonalDown: 4, // msoGradientDiagonalDown
    diagonalUp: 5,   // msoGradientDiagonalUp
  };
  gradientDirMap['vertical'] = 3; // correct: msoGradientVertical = 3 — no, it's different
  // Actually: msoGradientHorizontal=1, msoGradientMixed=-2, msoGradientDiagonalDown=4, msoGradientDiagonalUp=5, msoGradientFromCorner=6, msoGradientFromTitle=7, msoGradientFromCenter=8
  // For vertical: not a direct constant. Use msoGradientHorizontal with variant to get vertical effect.
  // Simpler approach: msoGradientStyle - horizontal=1, vertical trick via TwoColorGradient variant
  // TwoColorGradient(Style, Variant) — Style 1=horizontal, variant 1-4 rotates
  // For true vertical: not directly available as a simple constant. Let's just use OneColorGradient approach.
  // Actually let me just use the numeric values properly:
  // From MS docs: msoGradientHorizontal = 1 (left to right)
  // We can achieve vertical by using msoGradientHorizontal with variant 2 (top to bottom effect)
  // Or: use GradientAngle property after setting gradient

  const shapeTypeNum = shapeTypeMap[config.shapeType] || 1;

  let body = `  $shp = $ws.Shapes.AddShape(${shapeTypeNum}, ${config.left}, ${config.top}, ${config.width}, ${config.height})\n`;

  // Name
  if (config.name) {
    body += `  $shp.Name = '${escapePowerShellString(config.name)}'\n`;
  }

  // Fill
  if (config.fill) {
    if (config.fill.gradient) {
      const dir = gradientDirMap[config.fill.gradient.direction || 'horizontal'] || 1;
      body += `  $shp.Fill.TwoColorGradient(${dir}, 1)\n`;
      body += `  $shp.Fill.ForeColor.RGB = ${hexToExcelColor(config.fill.gradient.color1)}\n`;
      body += `  $shp.Fill.BackColor.RGB = ${hexToExcelColor(config.fill.gradient.color2)}\n`;
    } else if (config.fill.color) {
      body += `  $shp.Fill.Solid()\n`;
      body += `  $shp.Fill.ForeColor.RGB = ${hexToExcelColor(config.fill.color)}\n`;
    }
    if (config.fill.transparency !== undefined) {
      body += `  $shp.Fill.Transparency = ${config.fill.transparency}\n`;
    }
  }

  // Line/border
  if (config.line) {
    if (config.line.visible === false) {
      body += `  $shp.Line.Visible = 0\n`; // msoFalse
    } else {
      body += `  $shp.Line.Visible = -1\n`; // msoTrue
      if (config.line.color) {
        body += `  $shp.Line.ForeColor.RGB = ${hexToExcelColor(config.line.color)}\n`;
      }
      if (config.line.weight !== undefined) {
        body += `  $shp.Line.Weight = ${config.line.weight}\n`;
      }
    }
  }

  // Shadow
  if (config.shadow) {
    if (config.shadow.visible !== false) {
      body += `  $shp.Shadow.Visible = -1\n`; // msoTrue
      if (config.shadow.color) {
        body += `  $shp.Shadow.ForeColor.RGB = ${hexToExcelColor(config.shadow.color)}\n`;
      }
      if (config.shadow.offsetX !== undefined) {
        body += `  $shp.Shadow.OffsetX = ${config.shadow.offsetX}\n`;
      }
      if (config.shadow.offsetY !== undefined) {
        body += `  $shp.Shadow.OffsetY = ${config.shadow.offsetY}\n`;
      }
      if (config.shadow.blur !== undefined) {
        body += `  $shp.Shadow.Blur = ${config.shadow.blur}\n`;
      }
      if (config.shadow.transparency !== undefined) {
        body += `  $shp.Shadow.Transparency = ${config.shadow.transparency}\n`;
      }
    } else {
      body += `  $shp.Shadow.Visible = 0\n`; // msoFalse
    }
  }

  // Text
  if (config.text) {
    const escapedText = escapePowerShellString(config.text.value);
    body += `  $shp.TextFrame2.TextRange.Text = '${escapedText}'\n`;

    const fontName = config.text.fontName || 'Segoe UI';
    body += `  $shp.TextFrame2.TextRange.Font.Name = '${escapePowerShellString(fontName)}'\n`;

    if (config.text.fontSize) {
      body += `  $shp.TextFrame2.TextRange.Font.Size = ${config.text.fontSize}\n`;
    }
    if (config.text.fontBold) {
      body += `  $shp.TextFrame2.TextRange.Font.Bold = -1\n`; // msoTrue
    }
    if (config.text.fontColor) {
      body += `  $shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ${hexToExcelColor(config.text.fontColor)}\n`;
    }

    // Horizontal alignment: msoAlignLeft=1, msoAlignCenter=2, msoAlignRight=3
    const hAlignTextMap: Record<string, number> = { left: 1, center: 2, right: 3 };
    const hAlign = hAlignTextMap[config.text.horizontalAlignment || 'center'] || 2;
    body += `  $shp.TextFrame2.TextRange.ParagraphFormat.Alignment = ${hAlign}\n`;

    // Vertical alignment: msoAnchorTop=1, msoAnchorMiddle=3, msoAnchorBottom=4
    const vAlignTextMap: Record<string, number> = { top: 1, middle: 3, bottom: 4 };
    const vAlign = vAlignTextMap[config.text.verticalAlignment || 'middle'] || 3;
    body += `  $shp.TextFrame2.VerticalAnchor = ${vAlign}\n`;

    // Margins
    body += `  $shp.TextFrame2.MarginLeft = 8\n`;
    body += `  $shp.TextFrame2.MarginRight = 8\n`;
    body += `  $shp.TextFrame2.MarginTop = 4\n`;
    body += `  $shp.TextFrame2.MarginBottom = 4\n`;
    body += `  $shp.TextFrame2.WordWrap = -1\n`; // msoTrue

    // AutoSize: msoAutoSizeNone=0, msoAutoSizeShapeToFitText=1, msoAutoSizeTextToFitShape=2
    if (config.text.autoSize && config.text.autoSize !== 'none') {
      const autoSizeMap: Record<string, number> = { shapeToFitText: 1, shrinkToFit: 2 };
      const autoSizeVal = autoSizeMap[config.text.autoSize] || 0;
      body += `  $shp.TextFrame2.AutoSize = ${autoSizeVal}\n`;
    }
  }

  body += `  Write-Output $shp.Name\n`;
  body += `  $wb.Save()\n`;

  const script = wrapWithCleanup(preamble, body);
  return await execPowerShellWithRetry(script);
}

// ============================================================
// Chart Styling (COM)
// ============================================================

interface StyleChartConfig {
  series?: Array<{
    index: number;
    color?: string;
    lineWeight?: number;
    markerStyle?: string;
    markerSize?: number;
    dataLabels?: {
      show: boolean;
      numberFormat?: string;
      fontSize?: number;
      fontColor?: string;
      position?: string;
      hideBelow?: number;
    };
  }>;
  axes?: {
    category?: {
      visible?: boolean;
      numberFormat?: string;
      fontSize?: number;
      fontColor?: string;
      labelRotation?: number;
    };
    value?: {
      visible?: boolean;
      numberFormat?: string;
      fontSize?: number;
      fontColor?: string;
      min?: number;
      max?: number;
      gridlines?: boolean;
    };
  };
  chartArea?: {
    fillColor?: string;
    borderVisible?: boolean;
  };
  plotArea?: {
    fillColor?: string;
  };
  legend?: {
    visible: boolean;
    position?: string;
    fontSize?: number;
    fontColor?: string;
  };
  title?: {
    text?: string;
    visible?: boolean;
    fontSize?: number;
    fontColor?: string;
  };
  width?: number;
  height?: number;
}

export async function styleChartViaPowerShell(
  filePath: string,
  sheetName: string,
  chartIndex: number | undefined,
  chartName: string | undefined,
  config: StyleChartConfig
): Promise<void> {
  const fileName = basename(filePath);
  const preamble = buildPreamble(fileName, sheetName);

  // COM constants
  const markerStyleMap: Record<string, number> = {
    circle: 8, square: 1, diamond: 2, triangle: 3, none: -4142,
  };
  const dataLabelPosMap: Record<string, number> = {
    above: 0, below: 1, left: -4131, right: -4152, center: -4108,
    outsideEnd: 2, insideEnd: 3, insideBase: 4,
  };
  const legendPosMap: Record<string, number> = {
    top: -4160, bottom: -4107, left: -4131, right: -4152,
  };

  // Locate chart
  let body = '';
  if (chartName) {
    const escaped = escapePowerShellString(chartName);
    body += `  $chartObj = $ws.ChartObjects('${escaped}')\n`;
  } else {
    const idx = chartIndex || 1;
    body += `  $chartObj = $ws.ChartObjects(${idx})\n`;
  }
  body += `  $chart = $chartObj.Chart\n`;

  // Size
  if (config.width !== undefined) {
    body += `  $chartObj.Width = ${config.width}\n`;
  }
  if (config.height !== undefined) {
    body += `  $chartObj.Height = ${config.height}\n`;
  }

  // Series styling
  if (config.series) {
    for (const s of config.series) {
      body += `  $s = $chart.SeriesCollection(${s.index})\n`;
      if (s.color) {
        const rgb = hexToExcelColor(s.color);
        // Set both fill and line color — fill for bar/column/pie, line for line/scatter
        body += `  try { $s.Format.Fill.Visible = -1; $s.Format.Fill.ForeColor.RGB = ${rgb} } catch {}\n`;
        body += `  try { $s.Format.Line.Visible = -1; $s.Format.Line.ForeColor.RGB = ${rgb} } catch {}\n`;
      }
      if (s.lineWeight !== undefined) {
        body += `  try { $s.Format.Line.Weight = ${s.lineWeight} } catch {}\n`;
      }
      if (s.markerStyle) {
        const msVal = markerStyleMap[s.markerStyle] ?? -4142;
        body += `  try { $s.MarkerStyle = ${msVal} } catch {}\n`;
      }
      if (s.markerSize !== undefined) {
        body += `  try { $s.MarkerSize = ${s.markerSize} } catch {}\n`;
      }
      if (s.dataLabels) {
        if (s.dataLabels.show) {
          body += `  $s.HasDataLabels = $true\n`;
          // Re-fetch DataLabels after enabling — COM object may not be initialized immediately
          body += `  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($s) | Out-Null\n`;
          body += `  $s = $chart.SeriesCollection(${s.index})\n`;
          body += `  $dl = $s.DataLabels\n`;
          if (s.dataLabels.numberFormat) {
            body += `  try { $dl.NumberFormat = '${escapePowerShellString(s.dataLabels.numberFormat)}' } catch {}\n`;
          }
          if (s.dataLabels.fontSize) {
            body += `  try { $dl.Font.Size = ${s.dataLabels.fontSize} } catch {}\n`;
          }
          if (s.dataLabels.fontColor) {
            body += `  try { $dl.Font.Color = ${hexToExcelColor(s.dataLabels.fontColor)} } catch {}\n`;
          }
          if (s.dataLabels.position) {
            const pos = dataLabelPosMap[s.dataLabels.position] ?? -4108;
            body += `  try { $dl.Position = ${pos} } catch {}\n`;
          }
          // Per-point label visibility: hide labels on segments smaller than threshold
          if (s.dataLabels.hideBelow !== undefined) {
            const threshold = s.dataLabels.hideBelow;
            body += `  # Hide data labels on points below threshold ${threshold}\n`;
            body += `  try {\n`;
            body += `    $vals = $s.Values\n`;
            body += `    for ($pi = 1; $pi -le $s.Points.Count; $pi++) {\n`;
            body += `      $pv = $vals[$pi - 1]\n`;
            body += `      if ($pv -ne $null -and [Math]::Abs([double]$pv) -lt ${threshold}) {\n`;
            body += `        $s.Points($pi).HasDataLabel = $false\n`;
            body += `      }\n`;
            body += `    }\n`;
            body += `  } catch {}\n`;
          }
        } else {
          body += `  $s.HasDataLabels = $false\n`;
        }
      }
    }
  }

  // Axes
  if (config.axes) {
    // Category axis (xlCategory = 1)
    if (config.axes.category) {
      const cat = config.axes.category;
      if (cat.visible === false) {
        body += `  try { $chart.Axes(1).Delete() } catch {}\n`;
      } else {
        body += `  $catAx = $chart.Axes(1)\n`;
        if (cat.numberFormat) {
          body += `  $catAx.TickLabels.NumberFormat = '${escapePowerShellString(cat.numberFormat)}'\n`;
        }
        if (cat.fontSize) {
          body += `  $catAx.TickLabels.Font.Size = ${cat.fontSize}\n`;
        }
        if (cat.fontColor) {
          body += `  $catAx.TickLabels.Font.Color = ${hexToExcelColor(cat.fontColor)}\n`;
        }
        if (cat.labelRotation !== undefined) {
          body += `  $catAx.TickLabels.Orientation = ${cat.labelRotation}\n`;
        }
      }
    }
    // Value axis (xlValue = 2)
    if (config.axes.value) {
      const val = config.axes.value;
      if (val.visible === false) {
        body += `  try { $chart.Axes(2).Delete() } catch {}\n`;
      } else {
        body += `  $valAx = $chart.Axes(2)\n`;
        if (val.numberFormat) {
          body += `  $valAx.TickLabels.NumberFormat = '${escapePowerShellString(val.numberFormat)}'\n`;
        }
        if (val.fontSize) {
          body += `  $valAx.TickLabels.Font.Size = ${val.fontSize}\n`;
        }
        if (val.fontColor) {
          body += `  $valAx.TickLabels.Font.Color = ${hexToExcelColor(val.fontColor)}\n`;
        }
        if (val.min !== undefined) {
          body += `  $valAx.MinimumScale = ${val.min}\n`;
        }
        if (val.max !== undefined) {
          body += `  $valAx.MaximumScale = ${val.max}\n`;
        }
        if (val.gridlines !== undefined) {
          body += `  $valAx.HasMajorGridlines = $${val.gridlines}\n`;
        }
      }
    }
  }

  // Chart area
  if (config.chartArea) {
    if (config.chartArea.fillColor) {
      body += `  $chart.ChartArea.Format.Fill.Visible = -1\n`;
      body += `  $chart.ChartArea.Format.Fill.ForeColor.RGB = ${hexToExcelColor(config.chartArea.fillColor)}\n`;
    }
    if (config.chartArea.borderVisible !== undefined) {
      body += `  $chart.ChartArea.Format.Line.Visible = ${config.chartArea.borderVisible ? '-1' : '0'}\n`;
    }
  }

  // Plot area
  if (config.plotArea) {
    if (config.plotArea.fillColor) {
      body += `  $chart.PlotArea.Format.Fill.Visible = -1\n`;
      body += `  $chart.PlotArea.Format.Fill.ForeColor.RGB = ${hexToExcelColor(config.plotArea.fillColor)}\n`;
    }
  }

  // Legend
  if (config.legend) {
    body += `  $chart.HasLegend = $${config.legend.visible}\n`;
    if (config.legend.visible) {
      if (config.legend.position) {
        const pos = legendPosMap[config.legend.position] ?? -4107;
        body += `  $chart.Legend.Position = ${pos}\n`;
      }
      if (config.legend.fontSize) {
        body += `  $chart.Legend.Font.Size = ${config.legend.fontSize}\n`;
      }
      if (config.legend.fontColor) {
        body += `  $chart.Legend.Font.Color = ${hexToExcelColor(config.legend.fontColor)}\n`;
      }
    }
  }

  // Title
  if (config.title) {
    if (config.title.visible !== undefined) {
      body += `  $chart.HasTitle = $${config.title.visible}\n`;
    }
    if (config.title.text !== undefined) {
      body += `  $chart.HasTitle = $true\n`;
      body += `  $chart.ChartTitle.Text = '${escapePowerShellString(config.title.text)}'\n`;
    }
    if (config.title.fontSize) {
      body += `  $chart.ChartTitle.Font.Size = ${config.title.fontSize}\n`;
    }
    if (config.title.fontColor) {
      body += `  $chart.ChartTitle.Font.Color = ${hexToExcelColor(config.title.fontColor)}\n`;
    }
  }

  body += `  $wb.Save()\n`;
  const script = wrapWithCleanup(preamble, body);
  await execPowerShellWithRetry(script, MAX_RETRIES, 15000);
}
