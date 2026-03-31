import { loadWorkbook, getSheet, saveWorkbook, columnLetterToNumber, formatDataAsTable } from './helpers.js';
import type { ResponseFormat } from '../types.js';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  insertRowsLive,
  insertColumnsLive,
  unmergeCellsLive,
  saveFileLive,
} from './excel-live.js';

export async function insertRows(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  console.error(`[insertRows] Starting operation for ${filePath}, sheet: ${sheetName}, startRow: ${startRow}, count: ${count}`);

  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    console.error(`[insertRows] File is open in Excel, using live editing path`);
    try {
      // Live editing path - direct manipulation of open file
      await insertRowsLive(filePath, sheetName, startRow, count);
      await saveFileLive(filePath);

      console.error(`[insertRows] Successfully inserted ${count} row(s) via live editing`);

      return JSON.stringify({
        success: true,
        message: `Inserted ${count} row(s) at row ${startRow}`,
        startRow,
        count,
        method: 'live',
        note: 'Changes visible immediately in Excel',
      }, null, 2);
    } catch (error: any) {
      console.error(`[insertRows] Live editing failed, falling back to ExcelJS:`, error.message);
      // Fall through to ExcelJS fallback
    }
  }

  console.error(`[insertRows] Using ExcelJS fallback path`);

  // ExcelJS fallback - file not open or AppleScript failed
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Insert rows using ExcelJS
  sheet.spliceRows(startRow, 0, ...Array(count).fill([]));

  await saveWorkbook(workbook, filePath, createBackup);

  console.error(`[insertRows] Successfully inserted ${count} row(s) via ExcelJS`);

  return JSON.stringify({
    success: true,
    message: `Inserted ${count} row(s) at row ${startRow}`,
    startRow,
    count,
  }, null, 2);
}

export async function insertColumns(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number,
  createBackup: boolean = false
): Promise<string> {
  console.error(`[insertColumns] Starting operation for ${filePath}, sheet: ${sheetName}, startColumn: ${startColumn}, count: ${count}`);

  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    console.error(`[insertColumns] File is open in Excel, using live editing path`);
    try {
      // Live editing path - direct manipulation of open file
      await insertColumnsLive(filePath, sheetName, startColumn, count);
      await saveFileLive(filePath);

      console.error(`[insertColumns] Successfully inserted ${count} column(s) via live editing`);

      return JSON.stringify({
        success: true,
        message: `Inserted ${count} column(s) at column ${startColumn}`,
        startColumn,
        count,
        method: 'live',
        note: 'Changes visible immediately in Excel',
      }, null, 2);
    } catch (error: any) {
      console.error(`[insertColumns] Live editing failed, falling back to ExcelJS:`, error.message);
      // Fall through to ExcelJS fallback
    }
  }

  console.error(`[insertColumns] Using ExcelJS fallback path`);

  // ExcelJS fallback - file not open or AppleScript failed
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const colNumber = typeof startColumn === 'string' ? columnLetterToNumber(startColumn) : startColumn;

  // Insert columns using ExcelJS
  sheet.spliceColumns(colNumber, 0, ...Array(count).fill([]));

  await saveWorkbook(workbook, filePath, createBackup);

  console.error(`[insertColumns] Successfully inserted ${count} column(s) via ExcelJS`);

  return JSON.stringify({
    success: true,
    message: `Inserted ${count} column(s) at column ${startColumn}`,
    startColumn,
    count,
  }, null, 2);
}

export async function unmergeCells(
  filePath: string,
  sheetName: string,
  range: string,
  createBackup: boolean = false
): Promise<string> {
  console.error(`[unmergeCells] Starting operation for ${filePath}, sheet: ${sheetName}, range: ${range}`);

  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    console.error(`[unmergeCells] File is open in Excel, using live editing path`);
    try {
      // Live editing path - direct manipulation of open file
      await unmergeCellsLive(filePath, sheetName, range);
      await saveFileLive(filePath);

      console.error(`[unmergeCells] Successfully unmerged cells via live editing`);

      return JSON.stringify({
        success: true,
        message: `Cells unmerged in range ${range}`,
        range,
        method: 'live',
        note: 'Changes visible immediately in Excel',
      }, null, 2);
    } catch (error: any) {
      console.error(`[unmergeCells] Live editing failed, falling back to ExcelJS:`, error.message);
      // Fall through to ExcelJS fallback
    }
  }

  console.error(`[unmergeCells] Using ExcelJS fallback path`);

  // ExcelJS fallback - file not open or AppleScript failed
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Unmerge cells
  sheet.unMergeCells(range);

  await saveWorkbook(workbook, filePath, createBackup);

  console.error(`[unmergeCells] Successfully unmerged cells via ExcelJS`);

  return JSON.stringify({
    success: true,
    message: `Cells unmerged in range ${range}`,
    range,
  }, null, 2);
}

export async function getMergedCells(
  filePath: string,
  sheetName: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Get all merged cells from the sheet model
  const mergedCells: string[] = [];

  if (sheet.model.merges) {
    sheet.model.merges.forEach((merge) => {
      mergedCells.push(merge);
    });
  }

  if (responseFormat === 'markdown') {
    let md = `# Merged Cells in ${sheetName}\n\n`;
    md += `**Total merged ranges**: ${mergedCells.length}\n\n`;

    if (mergedCells.length > 0) {
      md += '## Merged Ranges\n\n';
      const tableData = mergedCells.map((range, index) => [index + 1, range]);
      md += formatDataAsTable(tableData, ['#', 'Range']);
    } else {
      md += '*No merged cells found*\n';
    }

    return md;
  }

  return JSON.stringify({
    sheetName,
    mergedCellsCount: mergedCells.length,
    mergedRanges: mergedCells,
  }, null, 2);
}
