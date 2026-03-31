import ExcelJS from 'exceljs';
import { loadWorkbook, getSheet, saveWorkbook, parseRange, columnNumberToLetter } from './helpers.js';
const FALLBACK_WARNING = 'Excel is running but this file was not detected as open via COM. Changes written to file on disk. If the file IS open in Excel with unsaved changes, saving from Excel may overwrite these. Close Excel without saving, then reopen to see changes.';

import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  updateCellLive,
  saveFileLive,
  addRowLive,
  writeRangeLive,
  setFormulaLive,
} from './excel-live.js';

export async function writeWorkbook(
  filePath: string,
  sheetName: string,
  data: any[][],
  createBackup: boolean = false
): Promise<string> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);

  // Write data
  data.forEach((row, rowIndex) => {
    const excelRow = sheet.getRow(rowIndex + 1);
    row.forEach((value, colIndex) => {
      excelRow.getCell(colIndex + 1).value = value;
    });
    excelRow.commit();
  });

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Workbook created at ${filePath}`,
    sheetName,
    rowsWritten: data.length,
    columnsWritten: data[0]?.length || 0,
  }, null, 2);
}

export async function updateCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  value: any,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[updateCell] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  console.error(`[updateCell] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  if (fileOpen) {
    console.error(`[updateCell] Using live editing method for real-time collaboration`);
    // Use live editing for real-time collaboration
    await updateCellLive(filePath, sheetName, cellAddress, value);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Cell ${cellAddress} updated (via Excel)`,
      cellAddress,
      newValue: value,
      method: 'live',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[updateCell] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const cell = sheet.getCell(cellAddress);
    cell.value = value;

    await saveWorkbook(workbook, filePath, createBackup);

    const response: Record<string, any> = {
      success: true,
      message: `Cell ${cellAddress} updated`,
      cellAddress,
      newValue: value,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    };
    if (excelRunning) {
      response.warning = FALLBACK_WARNING;
      response.method_reason = 'excel_running_but_file_not_detected_open';
    }
    return JSON.stringify(response, null, 2);
  }
}

export async function writeRange(
  filePath: string,
  sheetName: string,
  range: string,
  data: any[][],
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[writeRange] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  console.error(`[writeRange] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  if (fileOpen) {
    console.error(`[writeRange] Using live editing method for real-time collaboration`);
    // Use live editing for real-time collaboration
    const { startRow, startCol } = parseRange(range);
    const startCell = `${columnNumberToLetter(startCol)}${startRow}`;
    await writeRangeLive(filePath, sheetName, startCell, data);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Range ${range} updated (via Excel)`,
      range,
      rowsWritten: data.length,
      columnsWritten: data[0]?.length || 0,
      method: 'live',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[writeRange] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const { startRow, startCol } = parseRange(range);

    data.forEach((row, rowIndex) => {
      const excelRow = sheet.getRow(startRow + rowIndex);
      row.forEach((value, colIndex) => {
        excelRow.getCell(startCol + colIndex).value = value;
      });
      excelRow.commit();
    });

    await saveWorkbook(workbook, filePath, createBackup);

    const response: Record<string, any> = {
      success: true,
      message: `Range ${range} updated`,
      range,
      rowsWritten: data.length,
      columnsWritten: data[0]?.length || 0,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    };
    if (excelRunning) {
      response.warning = FALLBACK_WARNING;
      response.method_reason = 'excel_running_but_file_not_detected_open';
    }
    return JSON.stringify(response, null, 2);
  }
}

export async function addRow(
  filePath: string,
  sheetName: string,
  data: any[],
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[addRow] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  console.error(`[addRow] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  if (fileOpen) {
    console.error(`[addRow] Using live editing method for real-time collaboration`);
    // Use live editing for real-time collaboration
    await addRowLive(filePath, sheetName, data);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Row added (via Excel)`,
      cellsWritten: data.length,
      method: 'live',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[addRow] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const newRow = sheet.addRow(data);
    newRow.commit();

    await saveWorkbook(workbook, filePath, createBackup);

    const response: Record<string, any> = {
      success: true,
      message: `Row added at position ${newRow.number}`,
      rowNumber: newRow.number,
      cellsWritten: data.length,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    };
    if (excelRunning) {
      response.warning = FALLBACK_WARNING;
      response.method_reason = 'excel_running_but_file_not_detected_open';
    }
    return JSON.stringify(response, null, 2);
  }
}

export async function setFormula(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  formula: string,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and has this file open
  console.error(`[setFormula] Checking execution method for ${filePath}`);
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  console.error(`[setFormula] Method selection: excelRunning=${excelRunning}, fileOpen=${fileOpen}`);

  // Remove leading = if present
  const cleanFormula = formula.startsWith('=') ? formula.substring(1) : formula;

  if (fileOpen) {
    console.error(`[setFormula] Using live editing method for real-time collaboration`);
    // Use live editing for real-time collaboration
    await setFormulaLive(filePath, sheetName, cellAddress, cleanFormula);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Formula set in cell ${cellAddress} (via Excel)`,
      cellAddress,
      formula: `=${cleanFormula}`,
      method: 'live',
      note: 'Changes are visible immediately in Excel'
    }, null, 2);
  } else {
    console.error(`[setFormula] Using ExcelJS method (file-based editing)`);
    // Fallback to ExcelJS for file-based editing
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const cell = sheet.getCell(cellAddress);
    cell.value = { formula: cleanFormula };

    await saveWorkbook(workbook, filePath, createBackup);

    const response: Record<string, any> = {
      success: true,
      message: `Formula set in cell ${cellAddress}`,
      cellAddress,
      formula: `=${cleanFormula}`,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    };
    if (excelRunning) {
      response.warning = FALLBACK_WARNING;
      response.method_reason = 'excel_running_but_file_not_detected_open';
    }
    return JSON.stringify(response, null, 2);
  }
}
