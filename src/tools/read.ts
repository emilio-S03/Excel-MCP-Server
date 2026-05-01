import { loadWorkbook, getSheet, cellValueToString, formatDataAsTable, parseRange } from './helpers.js';
import type { WorkbookInfo, SheetInfo, ResponseFormat } from '../types.js';

export async function readWorkbook(filePath: string, responseFormat: ResponseFormat = 'json'): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  const sheets: SheetInfo[] = [];
  workbook.eachSheet((worksheet) => {
    sheets.push({
      name: worksheet.name,
      rowCount: worksheet.rowCount,
      columnCount: worksheet.columnCount,
      state: worksheet.state,
    });
  });

  const info: WorkbookInfo = {
    sheets,
    creator: workbook.creator,
    created: workbook.created,
    modified: workbook.modified,
  };

  if (responseFormat === 'markdown') {
    let md = `# Workbook: ${filePath}\n\n`;
    md += `**Created**: ${info.created ? new Date(info.created).toLocaleString() : 'N/A'}\n`;
    md += `**Modified**: ${info.modified ? new Date(info.modified).toLocaleString() : 'N/A'}\n`;
    md += `**Creator**: ${info.creator || 'N/A'}\n\n`;
    md += `## Sheets (${sheets.length})\n\n`;
    for (const sheet of sheets) {
      md += `- **${sheet.name}**: ${sheet.rowCount} rows × ${sheet.columnCount} columns`;
      if (sheet.state && sheet.state !== 'visible') {
        md += ` (${sheet.state})`;
      }
      md += '\n';
    }
    return md;
  }

  return JSON.stringify(info, null, 2);
}

export interface ReadSheetOptions {
  range?: string;
  offset?: number;
  limit?: number;
  maxCells?: number;
  responseFormat?: ResponseFormat;
}

export async function readSheet(
  filePath: string,
  sheetName: string,
  rangeOrOptions?: string | ReadSheetOptions,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const opts: ReadSheetOptions =
    typeof rangeOrOptions === 'string' || rangeOrOptions === undefined
      ? { range: rangeOrOptions, responseFormat }
      : { responseFormat, ...rangeOrOptions };

  const offset = Math.max(0, opts.offset ?? 0);
  const limit = opts.limit && opts.limit > 0 ? opts.limit : undefined;
  const maxCells = opts.maxCells && opts.maxCells > 0 ? opts.maxCells : undefined;
  const fmt = opts.responseFormat ?? 'json';

  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  let data: any[][] = [];
  let totalRows = 0;
  let totalColumns = 0;
  let truncatedByCells = false;

  if (opts.range) {
    const { startRow, startCol, endRow, endCol } = parseRange(opts.range);
    totalRows = endRow - startRow + 1;
    totalColumns = endCol - startCol + 1;
    const sliceStartRow = startRow + offset;
    const sliceEndRow = limit ? Math.min(endRow, sliceStartRow + limit - 1) : endRow;
    let cellsCollected = 0;
    for (let row = sliceStartRow; row <= sliceEndRow; row++) {
      const rowData: any[] = [];
      for (let col = startCol; col <= endCol; col++) {
        rowData.push(sheet.getRow(row).getCell(col).value);
      }
      cellsCollected += rowData.length;
      data.push(rowData);
      if (maxCells && cellsCollected >= maxCells) {
        truncatedByCells = true;
        break;
      }
    }
  } else {
    const allRows: any[][] = [];
    sheet.eachRow((row) => {
      const rowData: any[] = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowData.push(cell.value);
      });
      allRows.push(rowData);
    });
    totalRows = allRows.length;
    totalColumns = allRows[0]?.length || 0;
    const sliced = limit ? allRows.slice(offset, offset + limit) : allRows.slice(offset);
    if (maxCells) {
      let cells = 0;
      for (const row of sliced) {
        cells += row.length;
        data.push(row);
        if (cells >= maxCells) {
          truncatedByCells = true;
          break;
        }
      }
    } else {
      data = sliced;
    }
  }

  const returnedRows = data.length;
  const consumedRows = offset + returnedRows;
  const hasMore = consumedRows < totalRows;
  const nextOffset = hasMore ? consumedRows : null;

  if (fmt === 'markdown') {
    let md = `# Sheet: ${sheetName}\n\n`;
    if (opts.range) md += `**Range**: ${opts.range}\n`;
    md += `**Returned rows**: ${returnedRows} (of ${totalRows} total, offset ${offset})\n`;
    md += `**Columns**: ${totalColumns}\n`;
    if (hasMore) md += `**Next offset**: ${nextOffset}\n`;
    if (truncatedByCells) md += `**Truncated**: hit maxCells limit\n`;
    md += '\n## Data Preview\n\n';
    md += formatDataAsTable(data.slice(0, 100));
    if (returnedRows > 100) md += `\n\n*Showing first 100 of ${returnedRows} returned rows*`;
    return md;
  }

  return JSON.stringify(
    {
      sheetName,
      range: opts.range,
      rows: data,
      rowCount: returnedRows,
      columnCount: totalColumns,
      totalRows,
      offset,
      hasMore,
      nextOffset,
      truncatedByCells,
    },
    null,
    2
  );
}

export async function readRange(
  filePath: string,
  sheetName: string,
  range: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  return readSheet(filePath, sheetName, range, responseFormat);
}

export async function getCell(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cell = sheet.getCell(cellAddress);
  const result = {
    address: cellAddress,
    value: cell.value,
    type: cell.type,
    formula: cell.formula,
    numFmt: cell.numFmt,
  };

  if (responseFormat === 'markdown') {
    let md = `# Cell: ${cellAddress}\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    md += `**Value**: ${cellValueToString(cell.value)}\n`;
    md += `**Type**: ${cell.type}\n`;
    if (cell.formula) {
      md += `**Formula**: =${cell.formula}\n`;
    }
    if (cell.numFmt) {
      md += `**Format**: ${cell.numFmt}\n`;
    }
    return md;
  }

  return JSON.stringify(result, null, 2);
}

export async function getFormula(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cell = sheet.getCell(cellAddress);
  const result = {
    address: cellAddress,
    formula: cell.formula || null,
    value: cell.value,
  };

  if (responseFormat === 'markdown') {
    let md = `# Formula: ${cellAddress}\n\n`;
    md += `**Sheet**: ${sheetName}\n`;
    if (cell.formula) {
      md += `**Formula**: =${cell.formula}\n`;
      md += `**Result**: ${cellValueToString(cell.value)}\n`;
    } else {
      md += `**No formula** (direct value: ${cellValueToString(cell.value)})\n`;
    }
    return md;
  }

  return JSON.stringify(result, null, 2);
}
