import { loadWorkbook, getSheet, cellValueToString, formatDataAsTable, parseRange, columnNumberToLetter } from './helpers.js';
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

// ============================================================
// excel_read_sheet_merged_aware — read with merged-cell awareness
// ============================================================
// File-mode, cross-platform. Same shape as excel_read_sheet but post-
// processes the cell array so merged regions don't return mysterious
// empty cells. With fillMerged: true (default) every cell in a merged
// range is populated with the top-left value. With includeMergedMetadata:
// true, returns a `mergedCells` array describing each merged region in
// the read area.

export interface ReadSheetMergedAwareOptions {
  range?: string;
  fillMerged?: boolean;
  includeMergedMetadata?: boolean;
  responseFormat?: ResponseFormat;
}

interface MergedRegion {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  topLeft: string;
  range: string;
}

function parseMergeRangeString(range: string): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
  // Accepts "A1:C1" or "A1" (single-cell pseudo-merges).
  const matchPair = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (matchPair) {
    return {
      startCol: lettersToNum(matchPair[1]),
      startRow: parseInt(matchPair[2]),
      endCol: lettersToNum(matchPair[3]),
      endRow: parseInt(matchPair[4]),
    };
  }
  const matchOne = range.match(/^([A-Z]+)(\d+)$/);
  if (matchOne) {
    const c = lettersToNum(matchOne[1]);
    const r = parseInt(matchOne[2]);
    return { startCol: c, startRow: r, endCol: c, endRow: r };
  }
  return null;
}

function lettersToNum(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n;
}

function rectsOverlap(
  a: { startRow: number; startCol: number; endRow: number; endCol: number },
  b: { startRow: number; startCol: number; endRow: number; endCol: number }
): boolean {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endCol < b.startCol ||
    a.startCol > b.endCol
  );
}

export async function readSheetMergedAware(
  filePath: string,
  sheetName: string,
  options: ReadSheetMergedAwareOptions = {}
): Promise<string> {
  const fillMerged = options.fillMerged !== false; // default true
  const includeMergedMetadata = options.includeMergedMetadata === true;
  const fmt = options.responseFormat ?? 'json';

  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  // Collect merged regions from sheet model (same source as getMergedCells).
  const allMerges: MergedRegion[] = [];
  const rawMerges = (sheet.model as any).merges as string[] | undefined;
  if (rawMerges) {
    for (const m of rawMerges) {
      const parsed = parseMergeRangeString(m);
      if (!parsed) continue;
      allMerges.push({
        ...parsed,
        topLeft: `${columnNumberToLetter(parsed.startCol)}${parsed.startRow}`,
        range: m,
      });
    }
  }

  // Determine the read window.
  let startRow: number, startCol: number, endRow: number, endCol: number;
  if (options.range) {
    const parsed = parseRange(options.range);
    startRow = parsed.startRow;
    startCol = parsed.startCol;
    endRow = parsed.endRow;
    endCol = parsed.endCol;
  } else {
    startRow = 1;
    startCol = 1;
    // Use sheet bounds — fall back to 1x1 if the sheet is empty.
    endRow = Math.max(1, sheet.rowCount || 1);
    endCol = Math.max(1, sheet.columnCount || 1);
  }

  // Build cell value matrix.
  const data: any[][] = [];
  for (let r = startRow; r <= endRow; r++) {
    const row: any[] = [];
    for (let c = startCol; c <= endCol; c++) {
      row.push(sheet.getRow(r).getCell(c).value);
    }
    data.push(row);
  }

  // Find merges that intersect the read window.
  const window = { startRow, startCol, endRow, endCol };
  const mergesInWindow = allMerges.filter((m) => rectsOverlap(m, window));

  // Fill merged-region cells with the top-left value.
  if (fillMerged) {
    for (const m of mergesInWindow) {
      // Top-left value comes from the actual cell (lives outside window if
      // the merge starts before the window; we still need the value).
      const topLeftValue = sheet.getRow(m.startRow).getCell(m.startCol).value;
      const ir0 = Math.max(m.startRow, startRow);
      const ir1 = Math.min(m.endRow, endRow);
      const ic0 = Math.max(m.startCol, startCol);
      const ic1 = Math.min(m.endCol, endCol);
      for (let r = ir0; r <= ir1; r++) {
        for (let c = ic0; c <= ic1; c++) {
          data[r - startRow][c - startCol] = topLeftValue;
        }
      }
    }
  }

  const totalRows = endRow - startRow + 1;
  const totalCols = endCol - startCol + 1;

  // Build merged metadata if requested.
  let mergedCellsMeta: Array<{ topLeft: string; range: string; value: any }> | undefined;
  if (includeMergedMetadata) {
    mergedCellsMeta = mergesInWindow.map((m) => ({
      topLeft: m.topLeft,
      range: m.range,
      value: sheet.getRow(m.startRow).getCell(m.startCol).value,
    }));
  }

  if (fmt === 'markdown') {
    let md = `# Sheet (merged-aware): ${sheetName}\n\n`;
    if (options.range) md += `**Range**: ${options.range}\n`;
    md += `**Rows**: ${totalRows}\n`;
    md += `**Columns**: ${totalCols}\n`;
    md += `**fillMerged**: ${fillMerged}\n`;
    md += `**Merged regions in window**: ${mergesInWindow.length}\n\n`;
    md += '## Data\n\n';
    md += formatDataAsTable(data.slice(0, 100));
    if (totalRows > 100) md += `\n\n*Showing first 100 of ${totalRows} rows*`;
    if (includeMergedMetadata && mergedCellsMeta && mergedCellsMeta.length > 0) {
      md += `\n\n## Merged Cells\n\n`;
      const tableData = mergedCellsMeta.map((m, i) => [
        i + 1,
        m.range,
        m.topLeft,
        cellValueToString(m.value),
      ]);
      md += formatDataAsTable(tableData, ['#', 'Range', 'Top-Left', 'Value']);
    }
    return md;
  }

  const out: any = {
    sheetName,
    range: options.range,
    rows: data,
    rowCount: totalRows,
    columnCount: totalCols,
    fillMerged,
    mergedRegionsInWindow: mergesInWindow.length,
  };
  if (includeMergedMetadata) {
    out.mergedCells = mergedCellsMeta;
  }
  return JSON.stringify(out, null, 2);
}
