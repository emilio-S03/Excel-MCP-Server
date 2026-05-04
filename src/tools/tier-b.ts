/**
 * Tier B diagnostic tools (file-mode, cross-platform).
 *
 *  - excel_dependency_graph         : full workbook formula dependency graph
 *  - excel_compare_sheets           : structural diff between two sheets (same or different files)
 *  - excel_validate_named_range_targets : audit named ranges for invalid targets
 *  - excel_get_calculation_chain    : read xl/calcChain.xml from the .xlsx zip
 *
 * All tools are read-only; none mutate the workbook.
 */
import { promises as fs } from 'node:fs';
import JSZip from 'jszip';
import {
  loadWorkbook,
  getSheet,
  ensureFilePathAllowed,
  columnNumberToLetter,
  columnLetterToNumber,
  cellValueToString,
} from './helpers.js';
import type { CellValue } from 'exceljs';

// Cell/range reference regex — same shape as audit.ts findCircularReferences.
// Matches: A1, AA10, $A$1, A1:B10, Sheet1!A1, 'My Sheet'!A1:B2.
const CELL_REF_REGEX =
  /(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?!?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?/g;

// =============================================================================
// shared helpers
// =============================================================================

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
 * Parse a formula and return all referenced single-cell addresses.
 * Sheet-qualified refs return as `Sheet!A1`. Unqualified return as just `A1`.
 * Ranges are expanded into individual cells, capped at `maxExpand`.
 */
function extractCellRefs(formula: string, maxExpand = 5000): string[] {
  const refs: string[] = [];
  let m: RegExpExecArray | null;
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
    try {
      const startCol = columnLetterToNumber(c1);
      const endCol = columnLetterToNumber(c2);
      const startRow = Math.min(r1, r2);
      const endRow = Math.max(r1, r2);
      const lowCol = Math.min(startCol, endCol);
      const highCol = Math.max(startCol, endCol);
      const cellCount = (endRow - startRow + 1) * (highCol - lowCol + 1);
      if (cellCount > maxExpand) {
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

// =============================================================================
// 1. excel_dependency_graph
// =============================================================================

interface DependencyNode {
  cell: string;       // bare address e.g. "C1"
  sheet: string;
  formula: string;
  refsTo: string[];   // qualified "Sheet!Addr"
  refsFrom: string[]; // qualified "Sheet!Addr"
}

export async function dependencyGraph(
  filePath: string,
  sheetName?: string,
  format: 'json' | 'mermaid' = 'json'
): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  // Optional sheet filter just narrows which formulas we *origin* from.
  const sheets = sheetName ? [getSheet(workbook, sheetName)] : workbook.worksheets;

  // adjacency: "Sheet!Addr" -> Set of qualified refs
  const refsTo = new Map<string, Set<string>>();
  const refsFrom = new Map<string, Set<string>>();
  const formulaByKey = new Map<string, { sheet: string; cell: string; formula: string }>();

  for (const sheet of sheets) {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const f = extractFormula(cell.value);
        if (!f) return;
        const key = `${sheet.name}!${cell.address}`;
        formulaByKey.set(key, { sheet: sheet.name, cell: cell.address, formula: f });

        const refs = extractCellRefs(f).map((r) =>
          r.includes('!') ? r : `${sheet.name}!${r}`
        );
        const dedupe = new Set(refs);
        refsTo.set(key, dedupe);
        for (const ref of dedupe) {
          if (!refsFrom.has(ref)) refsFrom.set(ref, new Set());
          refsFrom.get(ref)!.add(key);
        }
      });
    });
  }

  const nodes: DependencyNode[] = [];
  for (const [key, info] of formulaByKey) {
    nodes.push({
      cell: info.cell,
      sheet: info.sheet,
      formula: info.formula,
      refsTo: Array.from(refsTo.get(key) ?? []),
      refsFrom: Array.from(refsFrom.get(key) ?? []),
    });
  }

  // Cycle detection: a key is "cyclic" if walking refsTo leads back to itself.
  const cyclic: string[] = [];
  for (const startKey of refsTo.keys()) {
    if (hasCycle(startKey, refsTo)) {
      cyclic.push(startKey);
    }
  }

  let totalEdges = 0;
  for (const set of refsTo.values()) totalEdges += set.size;

  const result: any = {
    totalNodes: nodes.length,
    totalEdges,
    nodes,
    cyclic,
  };

  if (format === 'mermaid') {
    result.mermaid = buildMermaid(refsTo, formulaByKey, 100);
  }

  return JSON.stringify(result, null, 2);
}

function hasCycle(start: string, refsTo: Map<string, Set<string>>): boolean {
  const stack: string[] = [start];
  const seen = new Set<string>();
  while (stack.length) {
    const cur = stack.pop()!;
    const next = refsTo.get(cur);
    if (!next) continue;
    for (const n of next) {
      if (n === start) return true;
      if (seen.has(n)) continue;
      seen.add(n);
      stack.push(n);
    }
  }
  return false;
}

function buildMermaid(
  refsTo: Map<string, Set<string>>,
  formulaByKey: Map<string, { sheet: string; cell: string; formula: string }>,
  cap: number
): string {
  const lines: string[] = ['graph TD'];
  let edges = 0;
  outer: for (const [from, tos] of refsTo) {
    if (!formulaByKey.has(from)) continue;
    for (const to of tos) {
      if (edges >= cap) break outer;
      lines.push(`  ${mermaidId(from)}["${from}"] --> ${mermaidId(to)}["${to}"]`);
      edges++;
    }
  }
  if (edges >= cap) lines.push(`  %% truncated at ${cap} edges`);
  return lines.join('\n');
}

function mermaidId(key: string): string {
  return key.replace(/[^A-Za-z0-9]/g, '_');
}

// =============================================================================
// 2. excel_compare_sheets
// =============================================================================

interface CompareDiff {
  address: string;
  side: 'left-only' | 'right-only' | 'both-changed';
  leftValue?: any;
  rightValue?: any;
  leftFormula?: string;
  rightFormula?: string;
}

export async function compareSheets(
  leftFile: string,
  leftSheet: string,
  rightFile: string,
  rightSheet: string,
  options: { includeValues?: boolean; includeFormulas?: boolean } = {}
): Promise<string> {
  ensureFilePathAllowed(leftFile);
  ensureFilePathAllowed(rightFile);

  const includeValues = options.includeValues ?? true;
  const includeFormulas = options.includeFormulas ?? true;

  const leftWb = await loadWorkbook(leftFile);
  const rightWb = await loadWorkbook(rightFile);
  const left = getSheet(leftWb, leftSheet);
  const right = getSheet(rightWb, rightSheet);

  // Collect every populated cell into address-keyed maps.
  const leftMap = collectCells(left);
  const rightMap = collectCells(right);

  const allAddrs = new Set<string>([...leftMap.keys(), ...rightMap.keys()]);

  const differences: CompareDiff[] = [];
  let addedCells = 0;
  let removedCells = 0;
  let changedCells = 0;
  let formulasChanged = 0;
  let truncated = false;
  const CAP = 500;

  for (const addr of allAddrs) {
    const l = leftMap.get(addr);
    const r = rightMap.get(addr);

    if (!l && r) {
      addedCells++;
      if (differences.length < CAP) {
        const diff: CompareDiff = { address: addr, side: 'right-only' };
        if (includeValues) diff.rightValue = r.value;
        if (includeFormulas && r.formula) diff.rightFormula = r.formula;
        differences.push(diff);
      } else {
        truncated = true;
      }
      continue;
    }
    if (l && !r) {
      removedCells++;
      if (differences.length < CAP) {
        const diff: CompareDiff = { address: addr, side: 'left-only' };
        if (includeValues) diff.leftValue = l.value;
        if (includeFormulas && l.formula) diff.leftFormula = l.formula;
        differences.push(diff);
      } else {
        truncated = true;
      }
      continue;
    }
    if (l && r) {
      const valueDiff =
        cellValueToString(l.rawValue) !== cellValueToString(r.rawValue);
      const formulaDiff = (l.formula ?? null) !== (r.formula ?? null);
      if (!valueDiff && !formulaDiff) continue;

      changedCells++;
      if (formulaDiff) formulasChanged++;
      if (differences.length < CAP) {
        const diff: CompareDiff = { address: addr, side: 'both-changed' };
        if (includeValues) {
          diff.leftValue = l.value;
          diff.rightValue = r.value;
        }
        if (includeFormulas) {
          if (l.formula) diff.leftFormula = l.formula;
          if (r.formula) diff.rightFormula = r.formula;
        }
        differences.push(diff);
      } else {
        truncated = true;
      }
    }
  }

  return JSON.stringify(
    {
      summary: { addedCells, removedCells, changedCells, formulasChanged },
      differences,
      truncated,
    },
    null,
    2
  );
}

function collectCells(
  sheet: import('exceljs').Worksheet
): Map<string, { value: any; rawValue: CellValue; formula: string | null }> {
  const map = new Map<string, { value: any; rawValue: CellValue; formula: string | null }>();
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const f = extractFormula(cell.value);
      // Surface a friendly value: cached result for formulas, raw value otherwise.
      let display: any = cell.value;
      if (f && cell.value && typeof cell.value === 'object' && 'result' in cell.value) {
        display = (cell.value as any).result;
      }
      map.set(cell.address, {
        value: display,
        rawValue: cell.value,
        formula: f,
      });
    });
  });
  return map;
}

// =============================================================================
// 3. excel_validate_named_range_targets
// =============================================================================

interface InvalidName {
  name: string;
  formula: string;
  reason: string;
}

export async function validateNamedRangeTargets(filePath: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);

  const definedNames: any = (workbook as any).definedNames;
  const entries: Array<{ name: string; ranges: string[] }> = [];

  if (definedNames?.model && Array.isArray(definedNames.model)) {
    for (const entry of definedNames.model) {
      entries.push({
        name: entry.name,
        ranges: Array.isArray(entry.ranges) ? entry.ranges.map(String) : [],
      });
    }
  }

  const invalid: InvalidName[] = [];
  let validCount = 0;

  // Build a cheap sheet-name lookup (case-sensitive, mirrors Excel/ExcelJS lookup).
  const sheetNames = new Set(workbook.worksheets.map((w) => w.name));

  for (const e of entries) {
    if (e.ranges.length === 0) {
      invalid.push({
        name: e.name,
        formula: '',
        reason: 'Named range has no target ranges defined.',
      });
      continue;
    }
    const reasons: string[] = [];
    for (const raw of e.ranges) {
      const r = parseDefinedNameTarget(raw);
      if (!r) {
        reasons.push(`Could not parse target "${raw}".`);
        continue;
      }
      if (!sheetNames.has(r.sheet)) {
        reasons.push(`Sheet "${r.sheet}" does not exist (target: ${raw}).`);
        continue;
      }
      const sheet = workbook.getWorksheet(r.sheet)!;
      // Out-of-bounds check: Excel's hard limits.
      const EXCEL_MAX_ROWS = 1048576;
      const EXCEL_MAX_COLS = 16384;
      if (
        r.startRow < 1 || r.endRow < 1 ||
        r.startRow > EXCEL_MAX_ROWS || r.endRow > EXCEL_MAX_ROWS ||
        r.startCol < 1 || r.endCol < 1 ||
        r.startCol > EXCEL_MAX_COLS || r.endCol > EXCEL_MAX_COLS
      ) {
        reasons.push(`Target range "${raw}" is outside Excel's hard bounds.`);
        continue;
      }
      // Used-range check: target above sheet's used area is suspicious but legal.
      // We ONLY flag when the sheet itself has a smaller used region than the
      // start of the named range (i.e. the named range can never have data).
      if (sheet.rowCount > 0 && r.startRow > sheet.rowCount) {
        reasons.push(
          `Target "${raw}" starts at row ${r.startRow} but sheet "${r.sheet}" only has ${sheet.rowCount} rows of data.`
        );
        continue;
      }
      if (sheet.columnCount > 0 && r.startCol > sheet.columnCount) {
        reasons.push(
          `Target "${raw}" starts at column ${columnNumberToLetter(r.startCol)} but sheet "${r.sheet}" only has ${sheet.columnCount} columns of data.`
        );
        continue;
      }
    }
    if (reasons.length === 0) {
      validCount++;
    } else {
      invalid.push({
        name: e.name,
        formula: e.ranges.join(','),
        reason: reasons.join(' '),
      });
    }
  }

  return JSON.stringify(
    {
      totalNames: entries.length,
      validCount,
      invalidCount: invalid.length,
      invalid,
    },
    null,
    2
  );
}

interface ParsedTarget {
  sheet: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

function parseDefinedNameTarget(raw: string): ParsedTarget | null {
  // Examples:
  //   'Sheet 1'!$A$1:$B$10
  //   Sheet1!$A$1
  //   Sheet1!$A$1:$B$2
  // The sheet name may be quoted (with embedded apostrophes escaped as '').
  const trimmed = raw.trim();
  const bangIdx = lastBangOutsideQuotes(trimmed);
  if (bangIdx < 0) return null;
  let sheetPart = trimmed.slice(0, bangIdx);
  const cellPart = trimmed.slice(bangIdx + 1);
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'")) {
    sheetPart = sheetPart.slice(1, -1).replace(/''/g, "'");
  }
  // Cell part: $A$1[:$B$10]
  const m = cellPart.match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i);
  if (!m) return null;
  const c1 = columnLetterToNumber(m[1].toUpperCase());
  const r1 = parseInt(m[2], 10);
  const c2 = m[3] ? columnLetterToNumber(m[3].toUpperCase()) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) : r1;
  return {
    sheet: sheetPart,
    startCol: Math.min(c1, c2),
    endCol: Math.max(c1, c2),
    startRow: Math.min(r1, r2),
    endRow: Math.max(r1, r2),
  };
}

function lastBangOutsideQuotes(s: string): number {
  let inQuote = false;
  let lastIdx = -1;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'") {
      // Excel-style escape: '' inside a quoted name. Skip both.
      if (inQuote && s[i + 1] === "'") {
        i++;
        continue;
      }
      inQuote = !inQuote;
      continue;
    }
    if (ch === '!' && !inQuote) lastIdx = i;
  }
  return lastIdx;
}

// =============================================================================
// 4. excel_get_calculation_chain
// =============================================================================

interface CalcChainEntry {
  cell: string;
  sheetId: number | null;
  sheetName: string | null;
}

export async function getCalculationChain(filePath: string): Promise<string> {
  ensureFilePathAllowed(filePath);

  const buf = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);
  const calcChainFile = zip.file('xl/calcChain.xml');

  if (!calcChainFile) {
    return JSON.stringify(
      {
        totalEntries: 0,
        chain: [],
        note: "No calcChain.xml present — Excel hasn't recalculated this file yet",
      },
      null,
      2
    );
  }

  const xml = await calcChainFile.async('string');

  // Build a sheetId -> sheetName lookup from xl/workbook.xml.
  const sheetIdToName = await buildSheetIdMap(zip);

  const entries: CalcChainEntry[] = [];
  // Each entry looks like <c r="A1" i="1" .../>. The `i` attribute carries the
  // sheet id; it can be omitted when it equals the previous entry's i.
  const cTagRe = /<c\b([^/>]*)\/?\s*>/g;
  let lastSheetId: number | null = null;
  let m: RegExpExecArray | null;
  while ((m = cTagRe.exec(xml)) !== null) {
    const attrs = m[1];
    const rMatch = attrs.match(/\br\s*=\s*"([^"]+)"/);
    const iMatch = attrs.match(/\bi\s*=\s*"(-?\d+)"/);
    if (!rMatch) continue;
    const cell = rMatch[1];
    let sheetId: number | null = lastSheetId;
    if (iMatch) {
      sheetId = parseInt(iMatch[1], 10);
      lastSheetId = sheetId;
    }
    const sheetName = sheetId !== null ? sheetIdToName.get(sheetId) ?? null : null;
    entries.push({ cell, sheetId, sheetName });
  }

  return JSON.stringify({ totalEntries: entries.length, chain: entries }, null, 2);
}

async function buildSheetIdMap(zip: JSZip): Promise<Map<number, string>> {
  const map = new Map<number, string>();
  const wbFile = zip.file('xl/workbook.xml');
  if (!wbFile) return map;
  const xml = await wbFile.async('string');
  // <sheet name="Foo" sheetId="1" r:id="rId1"/>
  const re = /<sheet\b([^/>]*)\/?\s*>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1];
    const nameMatch = attrs.match(/\bname\s*=\s*"([^"]*)"/);
    const idMatch = attrs.match(/\bsheetId\s*=\s*"(\d+)"/);
    if (nameMatch && idMatch) {
      map.set(parseInt(idMatch[1], 10), nameMatch[1]);
    }
  }
  return map;
}
