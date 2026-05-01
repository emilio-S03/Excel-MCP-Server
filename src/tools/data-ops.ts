/**
 * v3.1 — data operation tools: sort, autofilter, dedupe, paste-special.
 * All file-mode (cross-platform).
 */
import {
  loadWorkbook,
  getSheet,
  saveWorkbook,
  parseRange,
  cellValueToString,
  columnLetterToNumber,
} from './helpers.js';

function resolveColumnIndex(col: string | number, startCol: number): number {
  if (typeof col === 'number') return col;
  return columnLetterToNumber(col.toUpperCase()) - startCol + 1;
}

function compareCells(a: any, b: any): number {
  const av = a == null ? '' : a;
  const bv = b == null ? '' : b;
  if (typeof av === 'number' && typeof bv === 'number') return av - bv;
  const as = cellValueToString(av);
  const bs = cellValueToString(bv);
  if (as < bs) return -1;
  if (as > bs) return 1;
  return 0;
}

export async function sortRange(
  filePath: string,
  sheetName: string,
  range: string,
  options: {
    sortBy: Array<{ column: string | number; order?: 'asc' | 'desc' }>;
    hasHeader?: boolean;
    createBackup?: boolean;
  }
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const { startRow, startCol, endRow, endCol } = parseRange(range);

  // Read range into array
  const rows: any[][] = [];
  for (let r = startRow; r <= endRow; r++) {
    const row: any[] = [];
    for (let c = startCol; c <= endCol; c++) {
      row.push(sheet.getRow(r).getCell(c).value);
    }
    rows.push(row);
  }

  const headerRow = options.hasHeader ? rows.shift() : null;

  // Resolve sort columns to 1-based indices within the slice
  const keys = options.sortBy.map((k) => ({
    idx: resolveColumnIndex(k.column, startCol) - 1, // 0-based within row array
    order: k.order ?? 'asc',
  }));

  rows.sort((a, b) => {
    for (const k of keys) {
      const cmp = compareCells(a[k.idx], b[k.idx]);
      if (cmp !== 0) return k.order === 'desc' ? -cmp : cmp;
    }
    return 0;
  });

  // Write back
  const finalRows = headerRow ? [headerRow, ...rows] : rows;
  for (let i = 0; i < finalRows.length; i++) {
    const r = startRow + i;
    for (let c = startCol; c <= endCol; c++) {
      sheet.getRow(r).getCell(c).value = finalRows[i][c - startCol];
    }
  }

  await saveWorkbook(workbook, filePath, options.createBackup ?? false);
  return JSON.stringify({
    success: true,
    range,
    sortedRowCount: rows.length,
    keys: options.sortBy,
    hasHeader: !!options.hasHeader,
  }, null, 2);
}

export async function setAutoFilter(
  filePath: string,
  sheetName: string,
  range: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  sheet.autoFilter = range;
  await saveWorkbook(workbook, filePath, createBackup);
  return JSON.stringify({ success: true, sheetName, autoFilterRange: range }, null, 2);
}

export async function clearAutoFilter(
  filePath: string,
  sheetName: string,
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  (sheet as any).autoFilter = undefined;
  await saveWorkbook(workbook, filePath, createBackup);
  return JSON.stringify({ success: true, sheetName, cleared: true }, null, 2);
}

export async function removeDuplicates(
  filePath: string,
  sheetName: string,
  range: string,
  options: {
    columns?: Array<string | number>; // columns to consider for uniqueness; undefined = all
    hasHeader?: boolean;
    createBackup?: boolean;
  }
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const { startRow, startCol, endRow, endCol } = parseRange(range);

  const rows: any[][] = [];
  for (let r = startRow; r <= endRow; r++) {
    const row: any[] = [];
    for (let c = startCol; c <= endCol; c++) {
      row.push(sheet.getRow(r).getCell(c).value);
    }
    rows.push(row);
  }

  const headerRow = options.hasHeader ? rows.shift() : null;

  const checkCols: number[] = options.columns
    ? options.columns.map((c) => resolveColumnIndex(c, startCol) - 1)
    : Array.from({ length: endCol - startCol + 1 }, (_, i) => i);

  const seen = new Set<string>();
  const kept: any[][] = [];
  let removed = 0;
  for (const row of rows) {
    const key = checkCols.map((i) => cellValueToString(row[i])).join('\x1f');
    if (seen.has(key)) {
      removed++;
    } else {
      seen.add(key);
      kept.push(row);
    }
  }

  // Write deduped data back, then clear leftover rows
  const finalRows = headerRow ? [headerRow, ...kept] : kept;
  for (let i = 0; i < finalRows.length; i++) {
    const r = startRow + i;
    for (let c = startCol; c <= endCol; c++) {
      sheet.getRow(r).getCell(c).value = finalRows[i][c - startCol];
    }
  }
  // Clear the now-orphan rows at the bottom
  for (let r = startRow + finalRows.length; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      sheet.getRow(r).getCell(c).value = null;
    }
  }

  await saveWorkbook(workbook, filePath, options.createBackup ?? false);
  return JSON.stringify({
    success: true,
    range,
    rowsScanned: rows.length,
    duplicatesRemoved: removed,
    rowsKept: kept.length,
    keyColumns: checkCols.map((i) => i + 1),
  }, null, 2);
}

export async function pasteSpecial(
  filePath: string,
  sheetName: string,
  sourceRange: string,
  targetCell: string,
  options: {
    mode: 'values' | 'formulas' | 'formats' | 'transpose';
    createBackup?: boolean;
  }
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const { startRow, startCol, endRow, endCol } = parseRange(sourceRange);
  const tgtMatch = targetCell.match(/^([A-Z]+)(\d+)$/);
  if (!tgtMatch) throw new Error(`Invalid target cell: ${targetCell}`);
  const tgtCol = columnLetterToNumber(tgtMatch[1]);
  const tgtRow = parseInt(tgtMatch[2], 10);

  let cellsWritten = 0;
  const rows = endRow - startRow + 1;
  const cols = endCol - startCol + 1;

  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const src = sheet.getRow(startRow + i).getCell(startCol + j);
      const dstRow = options.mode === 'transpose' ? tgtRow + j : tgtRow + i;
      const dstCol = options.mode === 'transpose' ? tgtCol + i : tgtCol + j;
      const dst = sheet.getRow(dstRow).getCell(dstCol);

      switch (options.mode) {
        case 'values': {
          const v = src.value;
          if (v && typeof v === 'object' && 'result' in (v as any)) {
            dst.value = (v as any).result;
          } else if (v && typeof v === 'object' && 'text' in (v as any)) {
            dst.value = (v as any).text;
          } else {
            dst.value = v;
          }
          break;
        }
        case 'formulas': {
          dst.value = src.value;
          break;
        }
        case 'formats': {
          dst.style = JSON.parse(JSON.stringify(src.style ?? {}));
          break;
        }
        case 'transpose': {
          dst.value = src.value;
          break;
        }
      }
      cellsWritten++;
    }
  }

  await saveWorkbook(workbook, filePath, options.createBackup ?? false);
  return JSON.stringify({
    success: true,
    mode: options.mode,
    sourceRange,
    targetCell,
    cellsWritten,
  }, null, 2);
}
