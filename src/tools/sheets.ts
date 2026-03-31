import { loadWorkbook, getSheet, saveWorkbook } from './helpers.js';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  createSheetLive,
  deleteSheetLive,
  renameSheetLive,
  setSheetProtectionLive,
  saveFileLive,
} from './excel-live.js';

export async function createSheet(
  filePath: string,
  sheetName: string,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and if the file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Check if sheet already exists via ExcelJS (need to load workbook for validation)
    const workbook = await loadWorkbook(filePath);
    if (workbook.getWorksheet(sheetName)) {
      throw new Error(`Sheet "${sheetName}" already exists`);
    }

    // Live editing path
    await createSheetLive(filePath, sheetName);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Sheet "${sheetName}" created`,
      sheetName,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    const workbook = await loadWorkbook(filePath);

    // Check if sheet already exists
    if (workbook.getWorksheet(sheetName)) {
      throw new Error(`Sheet "${sheetName}" already exists`);
    }

    workbook.addWorksheet(sheetName);
    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Sheet "${sheetName}" created`,
      sheetName,
      method: 'exceljs',
    }, null, 2);
  }
}

export async function deleteSheet(
  filePath: string,
  sheetName: string,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and if the file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Validate sheet exists via ExcelJS (need to load workbook for validation)
    const workbook = await loadWorkbook(filePath);
    getSheet(workbook, sheetName); // Throws if sheet doesn't exist

    // Live editing path
    await deleteSheetLive(filePath, sheetName);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Sheet "${sheetName}" deleted`,
      sheetName,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    workbook.removeWorksheet(sheet.id);
    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Sheet "${sheetName}" deleted`,
      sheetName,
      method: 'exceljs',
    }, null, 2);
  }
}

export async function renameSheet(
  filePath: string,
  oldName: string,
  newName: string,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and if the file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Check if new name already exists via ExcelJS (need to load workbook for validation)
    const workbook = await loadWorkbook(filePath);
    getSheet(workbook, oldName); // Throws if old sheet doesn't exist

    if (workbook.getWorksheet(newName)) {
      throw new Error(`Sheet "${newName}" already exists`);
    }

    // Live editing path
    await renameSheetLive(filePath, oldName, newName);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Sheet renamed from "${oldName}" to "${newName}"`,
      oldName,
      newName,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, oldName);

    // Check if new name already exists
    if (workbook.getWorksheet(newName)) {
      throw new Error(`Sheet "${newName}" already exists`);
    }

    sheet.name = newName;
    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Sheet renamed from "${oldName}" to "${newName}"`,
      oldName,
      newName,
      method: 'exceljs',
    }, null, 2);
  }
}

export async function duplicateSheet(
  filePath: string,
  sourceSheetName: string,
  newSheetName: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sourceSheet = getSheet(workbook, sourceSheetName);

  // Check if new name already exists
  if (workbook.getWorksheet(newSheetName)) {
    throw new Error(`Sheet "${newSheetName}" already exists`);
  }

  // Create new sheet
  const newSheet = workbook.addWorksheet(newSheetName);

  // Copy all data and formatting
  sourceSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const newRow = newSheet.getRow(rowNumber);
    newRow.height = row.height;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);

      // Copy value
      newCell.value = cell.value;

      // Copy formatting
      newCell.style = { ...cell.style };
    });

    newRow.commit();
  });

  // Copy column widths
  sourceSheet.columns.forEach((column, index) => {
    if (column && column.width) {
      const newColumn = newSheet.getColumn(index + 1);
      newColumn.width = column.width;
    }
  });

  // Copy merged cells
  if (sourceSheet.model.merges) {
    sourceSheet.model.merges.forEach((merge) => {
      newSheet.mergeCells(merge);
    });
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Sheet "${sourceSheetName}" duplicated as "${newSheetName}"`,
    sourceSheetName,
    newSheetName,
    rowsCopied: sourceSheet.rowCount,
  }, null, 2);
}

export async function setSheetProtection(
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
  },
  createBackup: boolean = false
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    await setSheetProtectionLive(filePath, sheetName, protect, password, options);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: protect ? `Sheet "${sheetName}" protected` : `Sheet "${sheetName}" unprotected`,
      sheetName,
      protect,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    if (protect) {
      const protectionOptions: any = {};
      if (options) {
        protectionOptions.insertRows = options.allowInsertRows || false;
        protectionOptions.insertColumns = options.allowInsertColumns || false;
        protectionOptions.deleteRows = options.allowDeleteRows || false;
        protectionOptions.deleteColumns = options.allowDeleteColumns || false;
        protectionOptions.sort = options.allowSort || false;
        protectionOptions.autoFilter = options.allowAutoFilter || false;
        protectionOptions.formatCells = options.allowFormatCells || false;
        protectionOptions.formatColumns = options.allowFormatColumns || false;
        protectionOptions.formatRows = options.allowFormatRows || false;
      }
      sheet.protect(password || '', protectionOptions);
    } else {
      sheet.unprotect();
    }

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: protect ? `Sheet "${sheetName}" protected` : `Sheet "${sheetName}" unprotected`,
      sheetName,
      protect,
      method: 'exceljs',
    }, null, 2);
  }
}
