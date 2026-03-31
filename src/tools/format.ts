import type { CellFormat } from '../types.js';
import { loadWorkbook, getSheet, saveWorkbook, columnLetterToNumber } from './helpers.js';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  formatCellLive,
  setColumnWidthLive,
  setRowHeightLive,
  mergeCellsLive,
  saveFileLive,
  batchFormatLive,
} from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

export async function formatCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  format: CellFormat,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Live editing path - format cell in real-time
    console.error(`[formatCell] Using live editing for real-time editing`);

    // Convert CellFormat to live editing format
    const liveFormat: {
      fontName?: string;
      fontSize?: number;
      fontBold?: boolean;
      fontItalic?: boolean;
      fontColor?: string;
      fillColor?: string;
      horizontalAlignment?: string;
      verticalAlignment?: string;
    } = {};

    if (format.font) {
      if (format.font.name) liveFormat.fontName = format.font.name;
      if (format.font.size) liveFormat.fontSize = format.font.size;
      if (format.font.bold !== undefined) liveFormat.fontBold = format.font.bold;
      if (format.font.italic !== undefined) liveFormat.fontItalic = format.font.italic;
      if (format.font.color) liveFormat.fontColor = format.font.color;
    }

    if (format.fill && format.fill.fgColor) {
      liveFormat.fillColor = format.fill.fgColor;
    }

    if (format.alignment) {
      if (format.alignment.horizontal) liveFormat.horizontalAlignment = format.alignment.horizontal;
      if (format.alignment.vertical) liveFormat.verticalAlignment = format.alignment.vertical;
    }

    await formatCellLive(filePath, sheetName, cellAddress, liveFormat);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Cell ${cellAddress} formatted`,
      cellAddress,
      appliedFormats: Object.keys(format),
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    console.error(`[formatCell] Using ExcelJS (Excel not running or file not open)`);

    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const cell = sheet.getCell(cellAddress);

    // Apply font formatting
    if (format.font) {
      cell.font = {
        ...cell.font,
        name: format.font.name,
        size: format.font.size,
        bold: format.font.bold,
        italic: format.font.italic,
        underline: format.font.underline,
        color: format.font.color ? { argb: format.font.color } : undefined,
      };
    }

    // Apply fill formatting
    if (format.fill) {
      cell.fill = {
        type: 'pattern',
        pattern: format.fill.pattern,
        fgColor: format.fill.fgColor ? { argb: format.fill.fgColor } : undefined,
        bgColor: format.fill.bgColor ? { argb: format.fill.bgColor } : undefined,
      };
    }

    // Apply alignment
    if (format.alignment) {
      cell.alignment = {
        ...cell.alignment,
        ...format.alignment,
      };
    }

    // Apply borders
    if (format.border) {
      const border: any = {};
      if (format.border.top) {
        border.top = {
          style: format.border.top.style,
          color: format.border.top.color ? { argb: format.border.top.color } : undefined,
        };
      }
      if (format.border.left) {
        border.left = {
          style: format.border.left.style,
          color: format.border.left.color ? { argb: format.border.left.color } : undefined,
        };
      }
      if (format.border.bottom) {
        border.bottom = {
          style: format.border.bottom.style,
          color: format.border.bottom.color ? { argb: format.border.bottom.color } : undefined,
        };
      }
      if (format.border.right) {
        border.right = {
          style: format.border.right.style,
          color: format.border.right.color ? { argb: format.border.right.color } : undefined,
        };
      }
      cell.border = border;
    }

    // Apply number format
    if (format.numFmt) {
      cell.numFmt = format.numFmt;
    }

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Cell ${cellAddress} formatted`,
      cellAddress,
      appliedFormats: Object.keys(format),
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    }, null, 2);
  }
}

export async function setColumnWidth(
  filePath: string,
  sheetName: string,
  column: string | number,
  width: number,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Live editing path - set column width in real-time
    console.error(`[setColumnWidth] Using live editing for real-time editing`);

    await setColumnWidthLive(filePath, sheetName, column, width);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Column ${column} width set to ${width}`,
      column,
      width,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    console.error(`[setColumnWidth] Using ExcelJS (Excel not running or file not open)`);

    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const colNumber = typeof column === 'string' ? columnLetterToNumber(column) : column;
    const col = sheet.getColumn(colNumber);
    col.width = width;

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Column ${column} width set to ${width}`,
      column,
      width,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    }, null, 2);
  }
}

export async function setRowHeight(
  filePath: string,
  sheetName: string,
  row: number,
  height: number,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Live editing path - set row height in real-time
    console.error(`[setRowHeight] Using live editing for real-time editing`);

    await setRowHeightLive(filePath, sheetName, row, height);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Row ${row} height set to ${height}`,
      row,
      height,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    console.error(`[setRowHeight] Using ExcelJS (Excel not running or file not open)`);

    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const excelRow = sheet.getRow(row);
    excelRow.height = height;
    excelRow.commit();

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Row ${row} height set to ${height}`,
      row,
      height,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    }, null, 2);
  }
}

export async function mergeCells(
  filePath: string,
  sheetName: string,
  range: string,
  createBackup: boolean = false
): Promise<string> {
  // Check if Excel is running and file is open
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    // Live editing path - merge cells in real-time
    console.error(`[mergeCells] Using live editing for real-time editing`);

    await mergeCellsLive(filePath, sheetName, range);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Cells merged in range ${range}`,
      range,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    // ExcelJS fallback
    console.error(`[mergeCells] Using ExcelJS (Excel not running or file not open)`);

    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    sheet.mergeCells(range);

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Cells merged in range ${range}`,
      range,
      method: 'exceljs',
      note: 'File updated. Open in Excel to see changes.',
    }, null, 2);
  }
}

export async function batchFormat(
  filePath: string,
  sheetName: string,
  operations: Array<{
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
  }>
): Promise<string> {
  // This tool requires Excel to be running — it uses COM for bulk operations
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  await batchFormatLive(filePath, sheetName, operations);

  return JSON.stringify({
    success: true,
    message: `Applied ${operations.length} formatting operations to sheet "${sheetName}"`,
    operationCount: operations.length,
    method: 'live',
    note: 'All formatting applied in a single batch. Changes visible immediately in Excel.',
  }, null, 2);
}
