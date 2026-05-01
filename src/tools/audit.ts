/**
 * Formula auditing & workbook health-check tools (file-mode, cross-platform).
 *
 * - excel_find_formula_errors: surfaces #DIV/0!, #REF!, #N/A, #VALUE!, #NAME?, #NULL!, #NUM! cells
 * - excel_find_circular_references: best-effort one-hop circular ref detection
 * - excel_workbook_stats: bulk stats useful for "why is this file 50MB" debugging
 * - excel_list_formulas: inventory all formulas on a sheet (with optional filter)
 * - excel_trace_precedents: best-effort one-level precedent extraction for a cell
 *
 * These intentionally do NOT recompute formula results — they read the saved
 * cached results that ExcelJS exposes (cell.value or cell.result).
 */
import { stat } from 'node:fs/promises';
import {
  loadWorkbook,
  getSheet,
  cellValueToString,
  parseRange,
  columnNumberToLetter,
} from './helpers.js';
import type { CellValue } from 'exceljs';

// Excel error sentinel pattern (post-evaluation strings ExcelJS may store).
const ERROR_VALUE_REGEX = /^#(REF|N\/A|VALUE|DIV\/0|NAME|NULL|NUM)!?\??$/;

// Cell/range reference regex for parsing formulas.
// Matches: A1, AA10, $A$1, A1:B10, Sheet1!A1, 'My Sheet'!A1:B2.
const CELL_REF_REGEX =
  /(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?!?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?/g;

interface FormulaError {
  sheet: string;
  cell: string;
  formula: string | null;
  errorType: string;
  errorValue: string;
}

interface CircularReference {
  sheet: string;
  cell: string;
  formula: string;
  referencedCells: string[];
}

interface ListedFormula {
  cell: string;
  formula: string;
  result?: any;
}

interface PrecedentRef {
  sheet: string;
  cell: string;
  value: any;
  formula?: string;
}

/**
 * Detect whether a cell's value (in any of its representations) is an Excel error.
 */
function detectErrorValue(value: CellValue): { errorType: string; errorValue: string } | null {
  if (value === null || value === undefined) return null;

  // ExcelJS error type
  if (typeof value === 'object') {
    if ('error' in value && (value as any).error) {
      const err = String((value as any).error);
      return { errorType: err, errorValue: err };
    }
    if ('result' in value && (value as any).result !== undefined) {
      const r = (value as any).result;
      if (typeof r === 'object' && r && 'error' in r && r.error) {
        const err = String(r.error);
        return { errorType: err, errorValue: err };
      }
      if (typeof r === 'string' && ERROR_VALUE_REGEX.test(r)) {
        return { errorType: r, errorValue: r };
      }
    }
  }

  if (typeof value === 'string' && ERROR_VALUE_REGEX.test(value)) {
    return { errorType: value, errorValue: value };
  }

  return null;
}

/**
 * Extract the formula string from a cell value, if present.
 */
function extractFormula(value: CellValue): string | null {
  if (value && typeof value === 'object' && 'formula' in value && (value as any).formula) {
    return String((value as any).formula);
  }
  if (value && typeof value === 'object' && 'sharedFormula' in value && (value as any).sharedFormula) {
    return String((value as any).sharedFormula);
  }
  return null;
}

/**
 * Walk every cell of every sheet (or one sheet) and collect Excel error cells.
 */
export async function findFormulaErrors(filePath: string, sheetName?: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const errors: FormulaError[] = [];

  const sheets = sheetName ? [getSheet(workbook, sheetName)] : workbook.worksheets;

  for (const sheet of sheets) {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const det = detectErrorValue(cell.value);
        if (det) {
          errors.push({
            sheet: sheet.name,
            cell: cell.address,
            formula: extractFormula(cell.value),
            errorType: det.errorType,
            errorValue: det.errorValue,
          });
        }
      });
    });
  }

  return JSON.stringify({ totalErrors: errors.length, errors }, null, 2);
}

/**
 * Parse a formula string and return all referenced cells (expanded to single-cell list).
 * Sheet-qualified refs return as `Sheet!A1`. Unqualified return as just `A1`.
 *
 * Caps expansion at `maxExpand` cells to avoid runaway memory on huge ranges.
 */
function extractCellRefs(formula: string, maxExpand = 5000): string[] {
  const refs: string[] = [];
  let m: RegExpExecArray | null;
  // Reset since the regex is global.
  CELL_REF_REGEX.lastIndex = 0;
  while ((m = CELL_REF_REGEX.exec(formula)) !== null) {
    const sheetQuoted = m[1];
    const sheetUnquoted = m[2];
    const sheet = sheetQuoted ?? sheetUnquoted ?? null;
    const c1 = m[3].toUpperCase();
    const r1 = parseInt(m[4], 10);
    const c2 = m[5] ? m[5].toUpperCase() : null;
    const r2 = m[6] ? parseInt(m[6], 10) : null;

    if (!c2 || r2 === null) {
      refs.push(sheet ? `${sheet}!${c1}${r1}` : `${c1}${r1}`);
      continue;
    }
    // Range — expand into individual cells (capped).
    try {
      const startCol = colLettersToNum(c1);
      const endCol = colLettersToNum(c2);
      const startRow = Math.min(r1, r2);
      const endRow = Math.max(r1, r2);
      const lowCol = Math.min(startCol, endCol);
      const highCol = Math.max(startCol, endCol);
      const cellCount = (endRow - startRow + 1) * (highCol - lowCol + 1);
      if (cellCount > maxExpand) {
        // Just emit the corners + a note token; don't blow memory.
        refs.push(sheet ? `${sheet}!${c1}${r1}` : `${c1}${r1}`);
        refs.push(sheet ? `${sheet}!${c2}${r2}` : `${c2}${r2}`);
        continue;
      }
      for (let r = startRow; r <= endRow; r++) {
        for (let c = lowCol; c <= highCol; c++) {
          const addr = `${columnNumberToLetter(c)}${r}`;
          refs.push(sheet ? `${sheet}!${addr}` : addr);
        }
      }
    } catch {
      refs.push(sheet ? `${sheet}!${c1}${r1}:${c2}${r2}` : `${c1}${r1}:${c2}${r2}`);
    }
  }
  return refs;
}

function colLettersToNum(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n;
}

/**
 * Best-effort circular reference detection: a cell is "circular" if it
 * directly references itself, OR if any cell it references (one hop)
 * references it back.
 *
 * Cross-sheet refs are honored when resolvable; bare refs are scoped to
 * the cell's own sheet.
 */
export async function findCircularReferences(filePath: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  // Build a global formula map: "Sheet!ADDR" -> { formula, refs }.
  type Entry = { formula: string; refs: string[]; sheet: string; cell: string };
  const formulas = new Map<string, Entry>();

  for (const sheet of workbook.worksheets) {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const f = extractFormula(cell.value);
        if (!f) return;
        const refs = extractCellRefs(f).map((r) =>
          r.includes('!') ? r : `${sheet.name}!${r}`
        );
        formulas.set(`${sheet.name}!${cell.address}`, {
          formula: f,
          refs,
          sheet: sheet.name,
          cell: cell.address,
        });
      });
    });
  }

  const circular: CircularReference[] = [];

  for (const [key, entry] of formulas) {
    // Direct self-reference?
    if (entry.refs.includes(key)) {
      circular.push({
        sheet: entry.sheet,
        cell: entry.cell,
        formula: entry.formula,
        referencedCells: entry.refs,
      });
      continue;
    }
    // One-hop transitive cycle?
    let foundCycle = false;
    for (const ref of entry.refs) {
      const target = formulas.get(ref);
      if (!target) continue;
      if (target.refs.includes(key)) {
        foundCycle = true;
        break;
      }
    }
    if (foundCycle) {
      circular.push({
        sheet: entry.sheet,
        cell: entry.cell,
        formula: entry.formula,
        referencedCells: entry.refs,
      });
    }
  }

  return JSON.stringify({ totalCircular: circular.length, references: circular }, null, 2);
}

/**
 * Workbook-wide stats: cell counts, formulas, merged cells, named ranges, file size, per-sheet breakdown.
 */
export async function workbookStats(filePath: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  let fileSizeBytes = 0;
  try {
    const s = await stat(filePath);
    fileSizeBytes = s.size;
  } catch {
    // best-effort
  }

  let totalCells = 0;
  let formulaCells = 0;
  let mergedCellRanges = 0;
  let conditionalFormats = 0;
  let dataValidations = 0;
  let hyperlinks = 0;
  let images = 0;
  let charts = 0;
  let tables = 0;

  const sheetStats: Array<{
    sheet: string;
    rowCount: number;
    columnCount: number;
    cellsUsed: number;
    formulasUsed: number;
    sizeContribution: number;
  }> = [];

  // Total payload string length across all sheets (proxy for size contribution).
  let grandPayload = 0;
  const perSheetPayload: Map<string, number> = new Map();

  for (const sheet of workbook.worksheets) {
    let sCells = 0;
    let sFormulas = 0;
    let sPayload = 0;
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        sCells++;
        const f = extractFormula(cell.value);
        if (f) sFormulas++;
        const repr = cellValueToString(cell.value);
        sPayload += repr ? repr.length : 0;
      });
    });

    totalCells += sCells;
    formulaCells += sFormulas;
    grandPayload += sPayload;
    perSheetPayload.set(sheet.name, sPayload);

    // ExcelJS exposes _merges as an object map; fallback to model.merges.
    const merges =
      (sheet as any)._merges ??
      (sheet as any).model?.merges ??
      null;
    if (merges) {
      if (Array.isArray(merges)) mergedCellRanges += merges.length;
      else mergedCellRanges += Object.keys(merges).length;
    }

    const cfs: any[] = (sheet as any).conditionalFormattings ?? [];
    conditionalFormats += cfs.length;

    const dvModel = (sheet as any).dataValidations?.model ?? {};
    dataValidations += Object.keys(dvModel).length;

    // Hyperlinks: count cells with .hyperlink (also in sheet.hyperlinks if present).
    let sheetLinkCount = 0;
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value as any;
        if (v && typeof v === 'object' && v.hyperlink) sheetLinkCount++;
      });
    });
    hyperlinks += sheetLinkCount;

    // Images per sheet (ExcelJS getImages())
    try {
      const imgs = typeof (sheet as any).getImages === 'function' ? (sheet as any).getImages() : [];
      if (Array.isArray(imgs)) images += imgs.length;
    } catch {
      /* ignore */
    }

    // Tables per sheet (ExcelJS tables map)
    try {
      const t = (sheet as any).tables;
      if (t && typeof t === 'object') {
        tables += Object.keys(t).length;
      }
    } catch {
      /* ignore */
    }

    sheetStats.push({
      sheet: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      cellsUsed: sCells,
      formulasUsed: sFormulas,
      sizeContribution: sPayload,
    });
  }

  // Charts live at workbook level in ExcelJS model — best-effort traverse.
  try {
    const wbModel: any = (workbook as any).model ?? {};
    if (Array.isArray(wbModel.media)) {
      // Already counted images per-sheet via getImages(); leave alone.
    }
    // Fallback: each worksheet model may carry charts under .charts.
    for (const ws of workbook.worksheets) {
      const wsModel: any = (ws as any).model ?? {};
      if (Array.isArray(wsModel.charts)) charts += wsModel.charts.length;
    }
  } catch {
    /* ignore */
  }

  let namedRanges = 0;
  try {
    const dn: any = (workbook as any).definedNames;
    if (dn?.model && Array.isArray(dn.model)) namedRanges = dn.model.length;
  } catch {
    /* ignore */
  }

  return JSON.stringify(
    {
      totalSheets: workbook.worksheets.length,
      totalCells,
      formulaCells,
      mergedCellRanges,
      namedRanges,
      conditionalFormats,
      dataValidations,
      hyperlinks,
      images,
      charts,
      tables,
      fileSizeBytes,
      sheetStats,
    },
    null,
    2
  );
}

/**
 * Inventory all formulas on a sheet with optional filter.
 */
export async function listFormulas(
  filePath: string,
  sheetName: string,
  options: { filter?: 'all' | 'array' | 'shared'; maxResults?: number } = {}
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const filter = options.filter ?? 'all';
  const maxResults = options.maxResults && options.maxResults > 0 ? options.maxResults : Infinity;

  const formulas: ListedFormula[] = [];
  let truncated = false;

  sheet.eachRow({ includeEmpty: false }, (row) => {
    if (formulas.length >= maxResults) {
      truncated = true;
      return;
    }
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (formulas.length >= maxResults) {
        truncated = true;
        return;
      }
      const v = cell.value as any;
      if (!v || typeof v !== 'object') return;

      const isFormula = !!v.formula;
      const isShared = !!v.sharedFormula;
      const isArray = v.formulaType === 'array' || !!v.array;

      if (filter === 'array' && !isArray) return;
      if (filter === 'shared' && !isShared) return;
      if (filter === 'all' && !isFormula && !isShared) return;

      const formula = v.formula ?? v.sharedFormula;
      if (!formula) return;

      const out: ListedFormula = { cell: cell.address, formula: String(formula) };
      if ('result' in v && v.result !== undefined) {
        out.result = v.result;
      }
      formulas.push(out);
    });
  });

  return JSON.stringify(
    {
      totalFormulas: formulas.length,
      truncated,
      filter,
      sheetName,
      formulas,
    },
    null,
    2
  );
}

/**
 * Trace one-level precedents of a cell. Returns each referenced cell with its current value
 * and (if formula) its formula string. Cross-sheet refs are resolved when possible.
 */
export async function tracePrecedents(
  filePath: string,
  sheetName: string,
  cellAddress: string
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const cell = sheet.getCell(cellAddress);

  const formula = extractFormula(cell.value);
  if (!formula) {
    return JSON.stringify(
      {
        cell: cellAddress,
        formula: null,
        directPrecedents: [],
        depth: 0,
        message: 'Cell has no formula — no precedents to trace.',
      },
      null,
      2
    );
  }

  const refs = extractCellRefs(formula);
  const seen = new Set<string>();
  const directPrecedents: PrecedentRef[] = [];

  for (const ref of refs) {
    let refSheetName = sheetName;
    let refAddr = ref;
    if (ref.includes('!')) {
      const idx = ref.indexOf('!');
      refSheetName = ref.slice(0, idx);
      refAddr = ref.slice(idx + 1);
    }
    const key = `${refSheetName}!${refAddr}`;
    if (seen.has(key)) continue;
    seen.add(key);

    const targetSheet = workbook.getWorksheet(refSheetName);
    if (!targetSheet) {
      directPrecedents.push({
        sheet: refSheetName,
        cell: refAddr,
        value: null,
        formula: undefined,
      });
      continue;
    }

    try {
      // parseRange handles A1:B2 form; bare addresses skip and just take cell.
      if (refAddr.includes(':')) {
        const r = parseRange(refAddr);
        for (let row = r.startRow; row <= r.endRow; row++) {
          for (let col = r.startCol; col <= r.endCol; col++) {
            const addr = `${columnNumberToLetter(col)}${row}`;
            const c = targetSheet.getCell(addr);
            const f = extractFormula(c.value);
            directPrecedents.push({
              sheet: refSheetName,
              cell: addr,
              value: f ? (c.value as any).result ?? null : c.value,
              ...(f ? { formula: f } : {}),
            });
          }
        }
      } else {
        const c = targetSheet.getCell(refAddr);
        const f = extractFormula(c.value);
        directPrecedents.push({
          sheet: refSheetName,
          cell: refAddr,
          value: f ? (c.value as any).result ?? null : c.value,
          ...(f ? { formula: f } : {}),
        });
      }
    } catch {
      directPrecedents.push({
        sheet: refSheetName,
        cell: refAddr,
        value: null,
      });
    }
  }

  return JSON.stringify(
    {
      cell: cellAddress,
      formula,
      directPrecedents,
      depth: 1,
    },
    null,
    2
  );
}
