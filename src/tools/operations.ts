import { loadWorkbook, getSheet, saveWorkbook, parseRange, columnLetterToNumber } from './helpers.js';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  deleteRowsLive,
  deleteColumnsLive,
  saveFileLive,
} from './excel-live.js';

export async function deleteRows(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  console.error(`[deleteRows] Starting operation: file="${filePath}", sheet="${sheetName}", startRow=${startRow}, count=${count}`);

  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Live editing path - file is open in Excel
    console.error(`[deleteRows] Using live editing (file is open in Excel)`);
    try {
      await deleteRowsLive(filePath, sheetName, startRow, count);
      await saveFileLive(filePath);

      console.error(`[deleteRows] Live editing operation completed successfully`);
      return JSON.stringify({
        success: true,
        message: `Deleted ${count} row(s) starting from row ${startRow}`,
        startRow,
        count,
        method: 'live',
        note: 'Changes visible immediately in open Excel file',
      }, null, 2);
    } catch (error: any) {
      console.error(`[deleteRows] Live editing failed, falling back to ExcelJS:`, error.message);
      // Fall through to ExcelJS fallback
    }
  }

  // ExcelJS fallback - file not open or AppleScript failed
  console.error(`[deleteRows] Using ExcelJS fallback`);
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  sheet.spliceRows(startRow, count);

  await saveWorkbook(workbook, filePath, createBackup);

  console.error(`[deleteRows] ExcelJS operation completed successfully`);
  return JSON.stringify({
    success: true,
    message: `Deleted ${count} row(s) starting from row ${startRow}`,
    startRow,
    count,
  }, null, 2);
}

export async function deleteColumns(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  console.error(`[deleteColumns] Starting operation: file="${filePath}", sheet="${sheetName}", startColumn=${startColumn}, count=${count}`);

  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Live editing path - file is open in Excel
    console.error(`[deleteColumns] Using live editing (file is open in Excel)`);
    try {
      // Pass startColumn directly (can be string or number - handler handles both)
      await deleteColumnsLive(filePath, sheetName, startColumn, count);
      await saveFileLive(filePath);

      console.error(`[deleteColumns] Live editing operation completed successfully`);
      return JSON.stringify({
        success: true,
        message: `Deleted ${count} column(s) starting from column ${startColumn}`,
        startColumn,
        count,
        method: 'live',
        note: 'Changes visible immediately in open Excel file',
      }, null, 2);
    } catch (error: any) {
      console.error(`[deleteColumns] Live editing failed, falling back to ExcelJS:`, error.message);
      // Fall through to ExcelJS fallback
    }
  }

  // ExcelJS fallback - file not open or AppleScript failed
  console.error(`[deleteColumns] Using ExcelJS fallback`);
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const colNumber = typeof startColumn === 'string' ? columnLetterToNumber(startColumn) : startColumn;
  sheet.spliceColumns(colNumber, count);

  await saveWorkbook(workbook, filePath, createBackup);

  console.error(`[deleteColumns] ExcelJS operation completed successfully`);
  return JSON.stringify({
    success: true,
    message: `Deleted ${count} column(s) starting from column ${startColumn}`,
    startColumn,
    count,
  }, null, 2);
}

export async function copyRange(
  filePath: string,
  sourceSheetName: string,
  sourceRange: string,
  targetSheetName: string,
  targetCell: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sourceSheet = getSheet(workbook, sourceSheetName);
  const targetSheet = getSheet(workbook, targetSheetName);

  const { startRow, startCol, endRow, endCol } = parseRange(sourceRange);
  const targetCellMatch = targetCell.match(/^([A-Z]+)(\d+)$/);

  if (!targetCellMatch) {
    throw new Error(`Invalid target cell address: ${targetCell}`);
  }

  const targetStartCol = columnLetterToNumber(targetCellMatch[1]);
  const targetStartRow = parseInt(targetCellMatch[2]);

  // Copy data and formatting
  for (let row = startRow; row <= endRow; row++) {
    const rowOffset = row - startRow;
    const targetRowNum = targetStartRow + rowOffset;
    const targetRow = targetSheet.getRow(targetRowNum);

    for (let col = startCol; col <= endCol; col++) {
      const colOffset = col - startCol;
      const targetColNum = targetStartCol + colOffset;

      const sourceCell = sourceSheet.getRow(row).getCell(col);
      const targetCellObj = targetRow.getCell(targetColNum);

      // Copy value
      targetCellObj.value = sourceCell.value;

      // Copy formatting
      targetCellObj.style = { ...sourceCell.style };
    }

    targetRow.commit();
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Range copied from ${sourceSheetName}!${sourceRange} to ${targetSheetName}!${targetCell}`,
    sourceSheet: sourceSheetName,
    sourceRange,
    targetSheet: targetSheetName,
    targetCell,
    rowsCopied: endRow - startRow + 1,
    columnsCopied: endCol - startCol + 1,
  }, null, 2);
}
