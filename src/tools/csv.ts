import { promises as fs } from 'fs';
import { parse } from 'csv-parse/sync';
import { stringify } from 'csv-stringify/sync';
import ExcelJS from 'exceljs';
import {
  loadWorkbook,
  getSheet,
  saveWorkbook,
  parseRange,
  cellValueToString,
  ensureFilePathAllowed,
} from './helpers.js';

export async function csvImport(
  csvPath: string,
  targetXlsx: string,
  options: {
    sheetName?: string;
    delimiter?: string;
    hasHeader?: boolean;
    createBackup?: boolean;
  }
): Promise<string> {
  ensureFilePathAllowed(csvPath);
  ensureFilePathAllowed(targetXlsx);

  const csvBuffer = await fs.readFile(csvPath);
  const records: unknown[][] = parse(csvBuffer, {
    delimiter: options.delimiter ?? ',',
    skip_empty_lines: true,
    relax_column_count: true,
  });

  let workbook: ExcelJS.Workbook;
  let isNew = false;
  try {
    await fs.access(targetXlsx);
    workbook = await loadWorkbook(targetXlsx);
  } catch {
    workbook = new ExcelJS.Workbook();
    isNew = true;
  }

  const sheetName = options.sheetName ?? 'Sheet1';
  let sheet = workbook.getWorksheet(sheetName);
  if (sheet) {
    workbook.removeWorksheet(sheet.id);
  }
  sheet = workbook.addWorksheet(sheetName);

  records.forEach((row, rowIdx) => {
    const excelRow = sheet!.getRow(rowIdx + 1);
    (row as unknown[]).forEach((value, colIdx) => {
      const cell = excelRow.getCell(colIdx + 1);
      if (typeof value === 'string') {
        const num = Number(value);
        if (value !== '' && !Number.isNaN(num) && Number.isFinite(num) && value.trim() === String(num)) {
          cell.value = num;
        } else {
          cell.value = value;
        }
      } else {
        cell.value = value as ExcelJS.CellValue;
      }
    });
    excelRow.commit();
  });

  if (options.hasHeader && sheet.rowCount > 0) {
    sheet.getRow(1).font = { bold: true };
  }

  await saveWorkbook(workbook, targetXlsx, options.createBackup ?? false);

  return JSON.stringify(
    {
      success: true,
      message: isNew
        ? `Created ${targetXlsx} with sheet ${sheetName}`
        : `Imported into existing workbook (sheet ${sheetName} replaced)`,
      csvPath,
      targetXlsx,
      sheetName,
      rowsImported: records.length,
      columnsImported: records[0]?.length ?? 0,
    },
    null,
    2
  );
}

export async function csvExport(
  filePath: string,
  sheetName: string,
  csvPath: string,
  options: { range?: string; delimiter?: string }
): Promise<string> {
  ensureFilePathAllowed(csvPath);

  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const data: string[][] = [];

  if (options.range) {
    const { startRow, startCol, endRow, endCol } = parseRange(options.range);
    for (let r = startRow; r <= endRow; r++) {
      const row: string[] = [];
      for (let c = startCol; c <= endCol; c++) {
        row.push(cellValueToString(sheet.getRow(r).getCell(c).value));
      }
      data.push(row);
    }
  } else {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      const rowData: string[] = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowData.push(cellValueToString(cell.value));
      });
      data.push(rowData);
    });
  }

  const csv: string = stringify(data, {
    delimiter: options.delimiter ?? ',',
    quoted_string: true,
  });

  await fs.writeFile(csvPath, csv, 'utf8');

  return JSON.stringify(
    {
      success: true,
      message: `Exported ${data.length} rows to ${csvPath}`,
      filePath,
      sheetName,
      csvPath,
      rowsExported: data.length,
    },
    null,
    2
  );
}
