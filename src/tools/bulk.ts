/**
 * Tier A bulk-operation tools (v3.3) — atomic multi-cell reads/writes via ExcelJS.
 * All file-mode (cross-platform). Validate-then-apply pattern keeps writes
 * transactional: a single bad entry rejects the whole batch before any
 * mutation hits disk.
 */
import { loadWorkbook, getSheet, saveWorkbook, parseRange, columnNumberToLetter } from './helpers.js';

function hasCellValue(cell: any): boolean {
  const v = cell?.value;
  if (v === null || v === undefined) return false;
  if (typeof v === 'string' && v === '') return false;
  return true;
}

function hasAnyStyle(cell: any): boolean {
  // ExcelJS lazy-creates style objects, so check for truly meaningful content
  if (cell.numFmt && cell.numFmt !== 'General') return true;
  const f = cell.font;
  if (f && (f.name || f.size !== undefined || f.bold || f.italic || f.underline || f.color)) return true;
  const fill = cell.fill;
  if (fill && fill.type) return true;
  const b = cell.border;
  if (b && (b.top || b.left || b.right || b.bottom)) return true;
  const a = cell.alignment;
  if (a && (a.horizontal || a.vertical || a.wrapText || a.indent || a.textRotation)) return true;
  return false;
}

/**
 * excel_get_cell_styles_bulk — read formatting for an entire range in one call.
 */
export async function getCellStylesBulk(
  filePath: string,
  sheetName: string,
  range: string,
  includeEmpty: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const { startCol, startRow, endCol, endRow } = parseRange(range);

  const cells: any[] = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const cell = sheet.getCell(r, c);
      const address = `${columnNumberToLetter(c)}${r}`;
      const hasValue = hasCellValue(cell);
      const hasStyle = hasAnyStyle(cell);

      if (!includeEmpty && !hasValue && !hasStyle) {
        continue;
      }

      cells.push({
        address,
        font: cell.font ?? null,
        fill: cell.fill ?? null,
        border: cell.border ?? null,
        alignment: cell.alignment ?? null,
        numFmt: cell.numFmt ?? null,
        hasValue,
      });
    }
  }

  return JSON.stringify({
    sheetName,
    range,
    includeEmpty,
    totalCells: cells.length,
    cells,
  }, null, 2);
}

/**
 * excel_batch_write_formulas — atomic bulk formula write.
 * Validates every entry first; throws before any write if anything looks bad.
 */
export async function batchWriteFormulas(
  filePath: string,
  sheetName: string,
  formulas: Array<{ cell: string; formula: string }>,
  createBackup: boolean = false
): Promise<string> {
  // Pre-validate everything BEFORE loading the workbook so errors are cheap.
  const cellPattern = /^[A-Z]+\d+$/;
  for (const [i, entry] of formulas.entries()) {
    if (!entry || typeof entry !== 'object') {
      throw new Error(`formulas[${i}]: must be an object {cell, formula}`);
    }
    if (typeof entry.cell !== 'string' || !cellPattern.test(entry.cell)) {
      throw new Error(`formulas[${i}].cell: invalid cell address "${entry.cell}" (expected e.g. A1)`);
    }
    if (typeof entry.formula !== 'string' || entry.formula.trim() === '') {
      throw new Error(`formulas[${i}].formula: must be a non-empty string`);
    }
    // Reject obviously broken formulas: unbalanced parens
    const opens = (entry.formula.match(/\(/g) || []).length;
    const closes = (entry.formula.match(/\)/g) || []).length;
    if (opens !== closes) {
      throw new Error(`formulas[${i}].formula: unbalanced parentheses in "${entry.formula}"`);
    }
  }

  // All entries valid — load once, apply all, save once.
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const written: Array<{ cell: string; formula: string }> = [];
  for (const entry of formulas) {
    const cleanFormula = entry.formula.startsWith('=') ? entry.formula.substring(1) : entry.formula;
    sheet.getCell(entry.cell).value = { formula: cleanFormula } as any;
    written.push({ cell: entry.cell, formula: `=${cleanFormula}` });
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    written: written.length,
    sheetName,
    formulas: written,
    method: 'exceljs',
  }, null, 2);
}

/**
 * excel_create_named_range_bulk — create N named ranges in one call.
 * Validates everything first (sheet existence + range syntax), then applies all.
 */
export async function createNamedRangeBulk(
  filePath: string,
  names: Array<{ name: string; sheetName: string; range: string }>,
  createBackup: boolean = false
): Promise<string> {
  // Pre-validate names + ranges syntactically.
  const namePattern = /^[A-Za-z_][A-Za-z0-9_.]*$/;
  const rangePattern = /^[A-Z]+\d+(:[A-Z]+\d+)?$/;
  for (const [i, entry] of names.entries()) {
    if (!entry || typeof entry !== 'object') {
      throw new Error(`names[${i}]: must be {name, sheetName, range}`);
    }
    if (typeof entry.name !== 'string' || !namePattern.test(entry.name)) {
      throw new Error(`names[${i}].name: invalid named-range identifier "${entry.name}" (must start with letter/underscore, alphanumeric+underscore+dot only)`);
    }
    if (typeof entry.sheetName !== 'string' || entry.sheetName.length === 0) {
      throw new Error(`names[${i}].sheetName: must be a non-empty string`);
    }
    if (typeof entry.range !== 'string' || !rangePattern.test(entry.range)) {
      throw new Error(`names[${i}].range: invalid range "${entry.range}" (expected e.g. A1 or A1:D10)`);
    }
  }

  const workbook = await loadWorkbook(filePath);

  // Verify every referenced sheet exists BEFORE we write anything.
  for (const [i, entry] of names.entries()) {
    if (!workbook.getWorksheet(entry.sheetName)) {
      throw new Error(`names[${i}]: sheet not found: "${entry.sheetName}"`);
    }
  }

  const dn = (workbook as any).definedNames;
  if (!dn || typeof dn.add !== 'function') {
    throw new Error('Workbook has no definedNames API (incompatible ExcelJS version).');
  }

  const created: Array<{ name: string; sheetName: string; range: string }> = [];
  for (const entry of names) {
    dn.add(`'${entry.sheetName}'!${entry.range}`, entry.name);
    created.push({ name: entry.name, sheetName: entry.sheetName, range: entry.range });
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    created: created.length,
    names: created,
    method: 'exceljs',
  }, null, 2);
}
