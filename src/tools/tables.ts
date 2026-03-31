import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';
import { isExcelRunningLive, isFileOpenInExcelLive, createTableLive, saveFileLive } from './excel-live.js';

export async function createTable(
  filePath: string,
  sheetName: string,
  range: string,
  tableName: string,
  tableStyle: string = 'TableStyleMedium2',
  showFirstColumn: boolean = false,
  showLastColumn: boolean = false,
  showRowStripes: boolean = true,
  showColumnStripes: boolean = false,
  createBackup: boolean = false
): Promise<string> {
  // Parse range for validation
  const { startRow, startCol, endRow, endCol } = parseRange(range);

  // Check if Excel is running and file is open — use COM if so
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    await createTableLive(filePath, sheetName, range, tableName, tableStyle);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Table "${tableName}" created via COM`,
      range,
      tableName,
      style: tableStyle,
      method: 'live',
      note: 'Native Excel table created. Visible immediately in Excel.',
    }, null, 2);
  }

  // ExcelJS fallback
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Read the headers from the first row
  const headerRow = sheet.getRow(startRow);
  const headers: any[] = [];
  for (let col = startCol; col <= endCol; col++) {
    const cell = headerRow.getCell(col);
    headers.push({ name: cell.value?.toString() || `Column${col}` });
  }

  sheet.addTable({
    name: tableName,
    ref: range,
    headerRow: true,
    totalsRow: false,
    style: {
      theme: tableStyle as any,
      showFirstColumn,
      showLastColumn,
      showRowStripes,
      showColumnStripes,
    },
    columns: headers,
    rows: [],
  });

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Table "${tableName}" created`,
    range,
    tableName,
    style: tableStyle,
    columns: headers.length,
    rows: endRow - startRow,
    method: 'exceljs',
  }, null, 2);
}
