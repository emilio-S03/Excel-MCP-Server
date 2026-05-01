/**
 * v3.1 — PDF export. Live mode (requires Excel running):
 *   - Windows: Excel.Application.ActiveWorkbook.ExportAsFixedFormat (COM)
 *   - macOS: AppleScript "save as PDF"
 *   - Linux: not supported (no Excel)
 *
 * Falls back to a clear error pointing the user at LibreOffice or an external
 * converter when Excel isn't available.
 */
import { exec } from 'child_process';
import { promisify } from 'util';
import { platform } from 'os';
import { basename } from 'path';
import { writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { ensureFilePathAllowed } from './helpers.js';

const execAsync = promisify(exec);
const IS_WIN = platform() === 'win32';
const IS_MAC = platform() === 'darwin';

function escapePsString(s: string): string {
  return s.replace(/'/g, "''");
}

function escapeAsString(s: string): string {
  return s.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

export async function exportPdf(
  filePath: string,
  outputPath: string,
  options: {
    sheetName?: string;
    range?: string;
    openAfterPublish?: boolean;
  }
): Promise<string> {
  ensureFilePathAllowed(outputPath);

  if (IS_WIN) {
    return exportPdfWindows(filePath, outputPath, options);
  }
  if (IS_MAC) {
    return exportPdfMac(filePath, outputPath);
  }
  throw new Error(
    'excel_export_pdf: PDF export requires Excel (Windows COM or macOS AppleScript). ' +
      'On Linux, convert via LibreOffice: `soffice --headless --convert-to pdf yourfile.xlsx`.'
  );
}

async function exportPdfWindows(
  filePath: string,
  outputPath: string,
  options: { sheetName?: string; range?: string; openAfterPublish?: boolean }
): Promise<string> {
  const fileName = basename(filePath);
  const escapedFile = escapePsString(fileName);
  const escapedOut = escapePsString(outputPath);
  const escapedSheet = options.sheetName ? escapePsString(options.sheetName) : '';
  const escapedRange = options.range ? escapePsString(options.range) : '';

  // xlTypePDF = 0; xlQualityStandard = 0; OpenAfterPublish = false
  const lines: string[] = [
    `$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')`,
    `$wb = $excel.Workbooks | Where-Object { $_.Name -eq '${escapedFile}' }`,
    `if (-not $wb) { throw "Workbook '${escapedFile}' not found in open Excel instance" }`,
  ];
  if (options.sheetName && options.range) {
    lines.push(`$ws = $wb.Worksheets.Item('${escapedSheet}')`);
    lines.push(`$ws.Range('${escapedRange}').ExportAsFixedFormat(0, '${escapedOut}', 0, $true, $true, [Type]::Missing, [Type]::Missing, ${options.openAfterPublish ? '$true' : '$false'})`);
  } else if (options.sheetName) {
    lines.push(`$ws = $wb.Worksheets.Item('${escapedSheet}')`);
    lines.push(`$ws.ExportAsFixedFormat(0, '${escapedOut}', 0, $true, $true, [Type]::Missing, [Type]::Missing, ${options.openAfterPublish ? '$true' : '$false'})`);
  } else {
    lines.push(`$wb.ExportAsFixedFormat(0, '${escapedOut}', 0, $true, $true, [Type]::Missing, [Type]::Missing, ${options.openAfterPublish ? '$true' : '$false'})`);
  }
  lines.push(`Write-Output 'OK'`);

  const script = lines.join('\n');
  const tmpFile = join(tmpdir(), `excel-mcp-pdf-${Date.now()}.ps1`);
  try {
    writeFileSync(tmpFile, script, 'utf8');
    await execAsync(
      `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "${tmpFile}"`,
      { timeout: 60000 }
    );
  } finally {
    try { unlinkSync(tmpFile); } catch {}
  }

  return JSON.stringify({
    success: true,
    filePath,
    outputPath,
    method: 'COM ExportAsFixedFormat',
    sheetName: options.sheetName,
    range: options.range,
  }, null, 2);
}

async function exportPdfMac(filePath: string, outputPath: string): Promise<string> {
  const fileName = basename(filePath);
  const script = [
    `tell application "Microsoft Excel"`,
    `  set wb to workbook "${escapeAsString(fileName)}"`,
    `  save wb in "${escapeAsString(outputPath)}" as PDF file format`,
    `end tell`,
  ].join('\n');
  const args = script
    .split('\n')
    .filter((l) => l.trim())
    .map((l) => `-e "${l.replace(/"/g, '\\"')}"`)
    .join(' ');
  await execAsync(`osascript ${args}`, { timeout: 30000 });

  return JSON.stringify({
    success: true,
    filePath,
    outputPath,
    method: 'AppleScript save as PDF',
  }, null, 2);
}
