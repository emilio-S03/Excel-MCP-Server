/**
 * v3.1 — hyperlink management.
 * All file-mode (cross-platform).
 */
import { loadWorkbook, getSheet, saveWorkbook, cellValueToString } from './helpers.js';

export async function addHyperlink(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  target: string,
  options: { text?: string; tooltip?: string; createBackup?: boolean }
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const cell = sheet.getCell(cellAddress);
  const displayText = options.text ?? cellValueToString(cell.value) ?? target;

  cell.value = {
    text: displayText,
    hyperlink: target,
    ...(options.tooltip ? { tooltip: options.tooltip } : {}),
  } as any;

  // Apply hyperlink-style formatting (blue, underlined) if no style yet
  cell.font = {
    ...(cell.font ?? {}),
    color: { argb: 'FF0563C1' },
    underline: true,
  };

  await saveWorkbook(workbook, filePath, options.createBackup ?? false);
  return JSON.stringify({
    success: true,
    cellAddress,
    target,
    text: displayText,
  }, null, 2);
}

export async function removeHyperlink(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  options: { keepText?: boolean; createBackup?: boolean }
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const cell = sheet.getCell(cellAddress);
  const v = cell.value as any;

  if (!v || typeof v !== 'object' || !v.hyperlink) {
    return JSON.stringify({ success: true, cellAddress, action: 'no-op (cell has no hyperlink)' }, null, 2);
  }

  if (options.keepText !== false && v.text !== undefined) {
    cell.value = v.text;
  } else {
    cell.value = null;
  }

  // Clear the link-style formatting
  if (cell.font?.underline || (cell.font?.color as any)?.argb === 'FF0563C1') {
    cell.font = { ...(cell.font ?? {}), underline: false, color: undefined as any };
  }

  await saveWorkbook(workbook, filePath, options.createBackup ?? false);
  return JSON.stringify({
    success: true,
    cellAddress,
    action: 'hyperlink removed',
    keptText: options.keepText !== false,
  }, null, 2);
}
