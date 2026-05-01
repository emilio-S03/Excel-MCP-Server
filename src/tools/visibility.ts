/**
 * v3.1 — sheet/row/column visibility tools.
 * All file-mode (cross-platform).
 */
import { loadWorkbook, getSheet, saveWorkbook, columnLetterToNumber } from './helpers.js';

export async function setSheetVisibility(
  filePath: string,
  sheetName: string,
  state: 'visible' | 'hidden' | 'veryHidden',
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Refuse to hide the only visible sheet — Excel rejects this anyway
  if (state !== 'visible') {
    const visibleCount = workbook.worksheets.filter(
      (w) => (w.state ?? 'visible') === 'visible' && w.name !== sheetName
    ).length;
    if (visibleCount === 0) {
      throw new Error(
        `Cannot hide ${sheetName}: at least one sheet must remain visible. Make another sheet visible first.`
      );
    }
  }

  sheet.state = state;
  await saveWorkbook(workbook, filePath, createBackup);
  return JSON.stringify({
    success: true,
    sheetName,
    state,
    note:
      state === 'veryHidden'
        ? 'veryHidden sheets are not shown in the Format > Hide & Unhide menu — only VBA or this tool can re-show them.'
        : undefined,
  }, null, 2);
}

export async function listSheetVisibility(filePath: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheets = workbook.worksheets.map((w) => ({
    name: w.name,
    state: w.state ?? 'visible',
    rowCount: w.rowCount,
    columnCount: w.columnCount,
  }));
  return JSON.stringify({
    totalSheets: sheets.length,
    visibleCount: sheets.filter((s) => s.state === 'visible').length,
    hiddenCount: sheets.filter((s) => s.state === 'hidden').length,
    veryHiddenCount: sheets.filter((s) => s.state === 'veryHidden').length,
    sheets,
  }, null, 2);
}

export async function hideRows(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number,
  hidden: boolean = true,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  for (let r = startRow; r < startRow + count; r++) {
    sheet.getRow(r).hidden = hidden;
  }
  await saveWorkbook(workbook, filePath, createBackup);
  return JSON.stringify({
    success: true,
    sheetName,
    startRow,
    count,
    action: hidden ? 'hidden' : 'unhidden',
  }, null, 2);
}

export async function hideColumns(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number,
  hidden: boolean = true,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const startNum =
    typeof startColumn === 'number' ? startColumn : columnLetterToNumber(startColumn.toUpperCase());
  for (let c = startNum; c < startNum + count; c++) {
    sheet.getColumn(c).hidden = hidden;
  }
  await saveWorkbook(workbook, filePath, createBackup);
  return JSON.stringify({
    success: true,
    sheetName,
    startColumn,
    count,
    action: hidden ? 'hidden' : 'unhidden',
  }, null, 2);
}
