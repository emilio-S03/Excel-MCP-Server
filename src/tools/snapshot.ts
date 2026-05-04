/**
 * Tier C — Workflow-shaping tools.
 *
 *   excel_snapshot_create   — copies a .xlsx to a sibling snapshot file
 *   excel_snapshot_diff     — compares all sheets between filePath and snapshot
 *   excel_snapshot_restore  — overwrites filePath with snapshot contents
 *   excel_transaction       — atomic batch executor (auto-rollback on failure)
 *   excel_diff_before_after — convenience: snapshot + run ops + diff
 *
 * Snapshots ALWAYS live in the same directory as the source file so
 * ensureFilePathAllowed catches them automatically — they cannot be created
 * outside the sandbox.
 */
import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import { dirname, basename, join, resolve } from 'path';
import { randomBytes } from 'crypto';

import {
  ensureFilePathAllowed,
  loadWorkbook,
  columnNumberToLetter,
} from './helpers.js';

import { updateCell, writeRange, addRow, setFormula } from './write.js';
import { formatCell, mergeCells } from './format.js';
import { createSheet, deleteSheet, renameSheet } from './sheets.js';
import { deleteRows, deleteColumns, copyRange } from './operations.js';
import {
  insertRows,
  insertColumns,
  unmergeCells,
} from './advanced.js';
import { applyConditionalFormat } from './conditional.js';
import { createNamedRange, deleteNamedRange } from './named-ranges.js';
import { setDataValidation } from './validation.js';
import {
  sortRange,
  removeDuplicates,
  pasteSpecial,
} from './data-ops.js';
import { addHyperlink, removeHyperlink } from './hyperlinks.js';

// ----------------------------------------------------------------------------
// Snapshot create
// ----------------------------------------------------------------------------

function genTimestamp(): string {
  // YYYYMMDDTHHmmssSSS — filesystem-safe, sortable.
  const now = new Date();
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  return (
    now.getUTCFullYear().toString() +
    pad(now.getUTCMonth() + 1) +
    pad(now.getUTCDate()) +
    'T' +
    pad(now.getUTCHours()) +
    pad(now.getUTCMinutes()) +
    pad(now.getUTCSeconds()) +
    pad(now.getUTCMilliseconds(), 3)
  );
}

function deriveSnapshotPath(filePath: string, snapshotId: string): string {
  const dir = dirname(filePath);
  const base = basename(filePath);
  // Strip a single trailing .xlsx/.xlsm to keep names readable; otherwise
  // just append.
  const stem = base.replace(/\.(xlsx|xlsm)$/i, '');
  return join(dir, `${stem}.snapshot.${snapshotId}.xlsx`);
}

export async function snapshotCreate(
  filePath: string,
  snapshotId?: string,
): Promise<string> {
  ensureFilePathAllowed(filePath);
  // Confirm source exists.
  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`Source file not found: ${filePath}`);
  }

  const id = snapshotId && snapshotId.length > 0 ? snapshotId : genTimestamp();
  const snapshotPath = deriveSnapshotPath(filePath, id);
  ensureFilePathAllowed(snapshotPath);

  await fs.copyFile(filePath, snapshotPath);

  const stat = await fs.stat(snapshotPath);
  return JSON.stringify(
    {
      success: true,
      snapshotPath,
      snapshotId: id,
      fileSize: stat.size,
      timestamp: new Date().toISOString(),
    },
    null,
    2,
  );
}

// ----------------------------------------------------------------------------
// Snapshot diff
// ----------------------------------------------------------------------------

interface CellDiff {
  sheet: string;
  cell: string;
  before: any;
  after: any;
}

function cellValueForDiff(value: any): any {
  if (value === null || value === undefined) return null;
  if (typeof value === 'object') {
    if ('formula' in value && value.formula) return `=${value.formula}`;
    if ('result' in value) return value.result;
    if ('text' in value) return value.text;
    if ('richText' in value) {
      try {
        return (value.richText as any[]).map((rt) => rt.text).join('');
      } catch {
        return JSON.stringify(value);
      }
    }
    return JSON.stringify(value);
  }
  return value;
}

function valuesEqual(a: any, b: any): boolean {
  const av = cellValueForDiff(a);
  const bv = cellValueForDiff(b);
  if (av === bv) return true;
  if (av === null || bv === null) return av === bv;
  // Numbers: tolerate floating-point near-equality.
  if (typeof av === 'number' && typeof bv === 'number') {
    if (Number.isNaN(av) && Number.isNaN(bv)) return true;
    return Math.abs(av - bv) < 1e-9;
  }
  // Dates: compare time value
  if (av instanceof Date && bv instanceof Date) {
    return av.getTime() === bv.getTime();
  }
  return String(av) === String(bv);
}

interface SheetExtent {
  rowCount: number;
  columnCount: number;
}

function sheetExtent(sheet: ExcelJS.Worksheet): SheetExtent {
  // ExcelJS rowCount/columnCount can include trailing empties; use actuals.
  const rowCount = sheet.actualRowCount ?? sheet.rowCount ?? 0;
  const columnCount = sheet.actualColumnCount ?? sheet.columnCount ?? 0;
  return { rowCount, columnCount };
}

function diffSheet(
  beforeSheet: ExcelJS.Worksheet,
  afterSheet: ExcelJS.Worksheet,
  sheetName: string,
  remainingCap: number,
): { changes: CellDiff[]; cellChanges: number } {
  const before = sheetExtent(beforeSheet);
  const after = sheetExtent(afterSheet);
  const maxRow = Math.max(before.rowCount, after.rowCount);
  const maxCol = Math.max(before.columnCount, after.columnCount);

  const changes: CellDiff[] = [];
  let cellChanges = 0;

  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const beforeVal = beforeSheet.getRow(r).getCell(c).value;
      const afterVal = afterSheet.getRow(r).getCell(c).value;
      if (!valuesEqual(beforeVal, afterVal)) {
        cellChanges++;
        if (changes.length < remainingCap) {
          changes.push({
            sheet: sheetName,
            cell: `${columnNumberToLetter(c)}${r}`,
            before: cellValueForDiff(beforeVal),
            after: cellValueForDiff(afterVal),
          });
        }
      }
    }
  }

  return { changes, cellChanges };
}

const DIFF_CAP = 500;

export async function snapshotDiff(
  filePath: string,
  snapshotPath: string,
): Promise<string> {
  ensureFilePathAllowed(filePath);
  ensureFilePathAllowed(snapshotPath);

  // Load both workbooks. The snapshot is the "before"; the live file is "after".
  const beforeWb = await loadWorkbook(snapshotPath);
  const afterWb = await loadWorkbook(filePath);

  const beforeSheets = new Map<string, ExcelJS.Worksheet>();
  beforeWb.eachSheet((s) => beforeSheets.set(s.name, s));
  const afterSheets = new Map<string, ExcelJS.Worksheet>();
  afterWb.eachSheet((s) => afterSheets.set(s.name, s));

  const allNames = new Set<string>([
    ...beforeSheets.keys(),
    ...afterSheets.keys(),
  ]);

  let sheetsAdded = 0;
  let sheetsRemoved = 0;
  let totalCellChanges = 0;
  const perSheet: Array<{
    sheetName: string;
    status: 'added' | 'removed' | 'changed' | 'identical';
    cellChanges: number;
  }> = [];
  const differences: CellDiff[] = [];

  for (const name of allNames) {
    const beforeSheet = beforeSheets.get(name);
    const afterSheet = afterSheets.get(name);

    if (!beforeSheet && afterSheet) {
      sheetsAdded++;
      // Count "all cells" in the new sheet as added cells.
      const ext = sheetExtent(afterSheet);
      let added = 0;
      const remaining = Math.max(0, DIFF_CAP - differences.length);
      for (let r = 1; r <= ext.rowCount; r++) {
        for (let c = 1; c <= ext.columnCount; c++) {
          const v = afterSheet.getRow(r).getCell(c).value;
          if (v === null || v === undefined || v === '') continue;
          added++;
          if (differences.length < DIFF_CAP) {
            differences.push({
              sheet: name,
              cell: `${columnNumberToLetter(c)}${r}`,
              before: null,
              after: cellValueForDiff(v),
            });
          }
          if (differences.length >= DIFF_CAP && added >= remaining) {
            // continue counting but stop pushing (cheap loop)
          }
        }
      }
      totalCellChanges += added;
      perSheet.push({ sheetName: name, status: 'added', cellChanges: added });
      continue;
    }

    if (beforeSheet && !afterSheet) {
      sheetsRemoved++;
      const ext = sheetExtent(beforeSheet);
      let removed = 0;
      for (let r = 1; r <= ext.rowCount; r++) {
        for (let c = 1; c <= ext.columnCount; c++) {
          const v = beforeSheet.getRow(r).getCell(c).value;
          if (v === null || v === undefined || v === '') continue;
          removed++;
          if (differences.length < DIFF_CAP) {
            differences.push({
              sheet: name,
              cell: `${columnNumberToLetter(c)}${r}`,
              before: cellValueForDiff(v),
              after: null,
            });
          }
        }
      }
      totalCellChanges += removed;
      perSheet.push({ sheetName: name, status: 'removed', cellChanges: removed });
      continue;
    }

    // Both present — compare cell-by-cell.
    const remaining = Math.max(0, DIFF_CAP - differences.length);
    const { changes, cellChanges } = diffSheet(
      beforeSheet!,
      afterSheet!,
      name,
      remaining,
    );
    totalCellChanges += cellChanges;
    if (cellChanges === 0) {
      perSheet.push({ sheetName: name, status: 'identical', cellChanges: 0 });
    } else {
      perSheet.push({ sheetName: name, status: 'changed', cellChanges });
      for (const c of changes) {
        if (differences.length < DIFF_CAP) differences.push(c);
      }
    }
  }

  return JSON.stringify(
    {
      summary: { sheetsAdded, sheetsRemoved, totalCellChanges },
      perSheet,
      differences,
      truncated: totalCellChanges > differences.length,
      diffCap: DIFF_CAP,
    },
    null,
    2,
  );
}

// ----------------------------------------------------------------------------
// Snapshot restore
// ----------------------------------------------------------------------------

export async function snapshotRestore(
  filePath: string,
  snapshotPath: string,
  createBackup: boolean = false,
): Promise<string> {
  ensureFilePathAllowed(filePath);
  ensureFilePathAllowed(snapshotPath);

  // Confirm snapshot exists.
  try {
    await fs.access(snapshotPath);
  } catch {
    throw new Error(`Snapshot not found: ${snapshotPath}`);
  }

  let backupPath: string | undefined;
  if (createBackup) {
    try {
      await fs.access(filePath);
      backupPath = `${filePath}.pre-restore-backup.xlsx`;
      ensureFilePathAllowed(backupPath);
      await fs.copyFile(filePath, backupPath);
    } catch {
      // Source doesn't exist — nothing to back up.
    }
  }

  await fs.copyFile(snapshotPath, filePath);

  return JSON.stringify(
    {
      success: true,
      restoredFrom: snapshotPath,
      restoredTo: filePath,
      backupPath: backupPath ?? null,
    },
    null,
    2,
  );
}

// ----------------------------------------------------------------------------
// Transaction safelist & executor
// ----------------------------------------------------------------------------

/**
 * Tools we'll execute inside a transaction. ALL must be file-mode (ExcelJS).
 * COM-only tools are excluded — they have side effects on a running Excel
 * process that can't be undone by overwriting the .xlsx file.
 *
 * Each entry is (validatedArgs) -> Promise<string>. We pre-validate args via
 * the schema BEFORE invocation so the executor can stay simple.
 */
type OpExecutor = (args: any) => Promise<string>;

export const TRANSACTION_SAFELIST: Record<string, OpExecutor> = {
  excel_update_cell: (a) =>
    updateCell(a.filePath, a.sheetName, a.cellAddress, a.value, a.createBackup),
  excel_write_range: (a) =>
    writeRange(a.filePath, a.sheetName, a.range, a.data, a.createBackup),
  excel_set_formula: (a) =>
    setFormula(a.filePath, a.sheetName, a.cellAddress, a.formula, a.createBackup),
  excel_format_cell: (a) =>
    formatCell(a.filePath, a.sheetName, a.cellAddress, a.format, a.createBackup),
  excel_add_row: (a) =>
    addRow(a.filePath, a.sheetName, a.data, a.createBackup),
  excel_create_sheet: (a) =>
    createSheet(a.filePath, a.sheetName, a.createBackup),
  excel_delete_sheet: (a) =>
    deleteSheet(a.filePath, a.sheetName, a.createBackup),
  excel_rename_sheet: (a) =>
    renameSheet(a.filePath, a.oldName, a.newName, a.createBackup),
  excel_delete_rows: (a) =>
    deleteRows(a.filePath, a.sheetName, a.startRow, a.count, a.createBackup),
  excel_delete_columns: (a) =>
    deleteColumns(a.filePath, a.sheetName, a.startColumn, a.count, a.createBackup),
  excel_insert_rows: (a) =>
    insertRows(a.filePath, a.sheetName, a.startRow, a.count, a.createBackup),
  excel_insert_columns: (a) =>
    insertColumns(a.filePath, a.sheetName, a.startColumn, a.count, a.createBackup),
  excel_merge_cells: (a) =>
    mergeCells(a.filePath, a.sheetName, a.range, a.createBackup),
  excel_unmerge_cells: (a) =>
    unmergeCells(a.filePath, a.sheetName, a.range, a.createBackup),
  excel_copy_range: (a) =>
    copyRange(
      a.filePath,
      a.sourceSheetName,
      a.sourceRange,
      a.targetSheetName,
      a.targetCell,
      a.createBackup,
    ),
  excel_add_hyperlink: (a) =>
    addHyperlink(a.filePath, a.sheetName, a.cellAddress, a.target, {
      text: a.text,
      tooltip: a.tooltip,
      createBackup: a.createBackup,
    }),
  excel_remove_hyperlink: (a) =>
    removeHyperlink(a.filePath, a.sheetName, a.cellAddress, {
      keepText: a.keepText,
      createBackup: a.createBackup,
    }),
  excel_sort: (a) =>
    sortRange(a.filePath, a.sheetName, a.range, {
      sortBy: a.sortBy,
      hasHeader: a.hasHeader,
      createBackup: a.createBackup,
    }),
  excel_remove_duplicates: (a) =>
    removeDuplicates(a.filePath, a.sheetName, a.range, {
      columns: a.columns,
      hasHeader: a.hasHeader,
      createBackup: a.createBackup,
    }),
  excel_paste_special: (a) =>
    pasteSpecial(a.filePath, a.sheetName, a.sourceRange, a.targetCell, {
      mode: a.mode,
      createBackup: a.createBackup,
    }),
  excel_create_named_range: (a) =>
    createNamedRange(a.filePath, a.name, a.sheetName, a.range, a.createBackup),
  excel_delete_named_range: (a) =>
    deleteNamedRange(a.filePath, a.name, a.createBackup),
  excel_set_data_validation: (a) =>
    setDataValidation(
      a.filePath,
      a.sheetName,
      a.range,
      a.validationType,
      a.formula1,
      a.operator,
      a.formula2,
      a.showErrorMessage,
      a.errorTitle,
      a.errorMessage,
      a.createBackup,
    ),
  excel_apply_conditional_format: (a) =>
    applyConditionalFormat(
      a.filePath,
      a.sheetName,
      a.range,
      a.ruleType,
      a.condition,
      a.style,
      a.colorScale,
      a.createBackup,
    ),
};

function tempSnapshotPath(filePath: string): string {
  const dir = dirname(filePath);
  const base = basename(filePath);
  const stem = base.replace(/\.(xlsx|xlsm)$/i, '');
  const rand = randomBytes(6).toString('hex');
  return join(dir, `${stem}.tx-snapshot.${rand}.xlsx`);
}

async function silentUnlink(path: string): Promise<void> {
  try {
    await fs.unlink(path);
  } catch {
    /* ignore */
  }
}

interface OpSpec {
  tool: string;
  args: Record<string, any>;
}

function rejectUnsafeTool(tool: string): never {
  const err = new Error(
    `excel_transaction: tool "${tool}" is not in the transaction safelist. ` +
      `Only file-mode write tools can be executed transactionally — COM-only, AppleScript, ` +
      `PowerShell, and live-mode tools are excluded because their side effects on a running ` +
      `Excel process can't be undone by file rollback. Safelist: ${Object.keys(
        TRANSACTION_SAFELIST,
      )
        .sort()
        .join(', ')}.`,
  );
  (err as any).code = 'PLATFORM_UNSUPPORTED';
  throw err;
}

interface ExecuteOptions {
  /** If true, snapshot and rollback on failure. */
  rollbackOnError: boolean;
  /** Optional pre-existing snapshot path (used by diff_before_after to share). */
  preCreatedSnapshot?: string;
  /** If true, leave the snapshot in place on success (caller takes ownership). */
  keepSnapshotOnSuccess: boolean;
}

interface ExecuteResult {
  snapshotPath: string;
  snapshotKept: boolean;
  results: any[];
}

async function executeOperations(
  filePath: string,
  operations: OpSpec[],
  opts: ExecuteOptions,
): Promise<{ ok: true; data: ExecuteResult } | {
  ok: false;
  failedAtStep: number;
  error: string;
  operationsCompleted: number;
  rolledBack: boolean;
  snapshotPath: string;
}> {
  ensureFilePathAllowed(filePath);

  // Validate all operation tool names BEFORE doing anything destructive.
  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    if (!op || typeof op.tool !== 'string') {
      throw new Error(
        `Operation ${i} is malformed: expected { tool, args }, got ${JSON.stringify(op)}`,
      );
    }
    if (!(op.tool in TRANSACTION_SAFELIST)) {
      rejectUnsafeTool(op.tool);
    }
  }

  // Confirm source exists.
  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`File not found: ${filePath}`);
  }

  const snapshotPath = opts.preCreatedSnapshot ?? tempSnapshotPath(filePath);
  ensureFilePathAllowed(snapshotPath);

  if (!opts.preCreatedSnapshot) {
    await fs.copyFile(filePath, snapshotPath);
  }

  const results: any[] = [];

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    const exec = TRANSACTION_SAFELIST[op.tool];
    // Force-set filePath on the args so callers can't smuggle a different
    // target into a transactional call.
    const args = { ...(op.args ?? {}), filePath };
    try {
      const raw = await exec(args);
      let parsed: any = raw;
      try {
        parsed = JSON.parse(raw);
      } catch {
        /* leave raw */
      }
      results.push({ step: i, tool: op.tool, result: parsed });
    } catch (err: any) {
      const errorMessage = err instanceof Error ? err.message : String(err);

      if (opts.rollbackOnError) {
        // Restore the snapshot, then delete it.
        try {
          await fs.copyFile(snapshotPath, filePath);
        } catch (restoreErr: any) {
          // If we can't restore, the snapshot is the user's recovery path.
          return {
            ok: false,
            failedAtStep: i,
            error: `${errorMessage} (and rollback FAILED: ${restoreErr?.message ?? restoreErr}; snapshot preserved at ${snapshotPath})`,
            operationsCompleted: i,
            rolledBack: false,
            snapshotPath,
          };
        }
        await silentUnlink(snapshotPath);
        return {
          ok: false,
          failedAtStep: i,
          error: errorMessage,
          operationsCompleted: i,
          rolledBack: true,
          snapshotPath,
        };
      }

      // Non-rollback path: stop on error but leave the file as-is and the
      // snapshot in place so caller can diagnose.
      return {
        ok: false,
        failedAtStep: i,
        error: errorMessage,
        operationsCompleted: i,
        rolledBack: false,
        snapshotPath,
      };
    }
  }

  // All ops succeeded.
  if (!opts.keepSnapshotOnSuccess) {
    await silentUnlink(snapshotPath);
  }

  return {
    ok: true,
    data: {
      snapshotPath,
      snapshotKept: opts.keepSnapshotOnSuccess,
      results,
    },
  };
}

export async function transaction(
  filePath: string,
  operations: OpSpec[],
  _createBackup: boolean = false,
): Promise<string> {
  // _createBackup is accepted for symmetry with other write tools but is
  // effectively redundant here — the snapshot IS a backup. We keep the param
  // in the schema so future expansion (e.g., copy snapshot to a permanent
  // .backup file on success) is non-breaking.
  void _createBackup;

  const result = await executeOperations(filePath, operations, {
    rollbackOnError: true,
    keepSnapshotOnSuccess: false,
  });

  if (!result.ok) {
    return JSON.stringify(
      {
        success: false,
        failedAtStep: result.failedAtStep,
        error: result.error,
        operationsCompleted: result.operationsCompleted,
        rolledBack: result.rolledBack,
      },
      null,
      2,
    );
  }

  return JSON.stringify(
    {
      success: true,
      operationsExecuted: result.data.results.length,
      results: result.data.results,
    },
    null,
    2,
  );
}

export async function diffBeforeAfter(
  filePath: string,
  operations: OpSpec[],
  keepSnapshot: boolean = false,
): Promise<string> {
  ensureFilePathAllowed(filePath);

  // Pre-create the snapshot so we have a stable "before" path to diff against.
  const snapshotPath = tempSnapshotPath(filePath);
  ensureFilePathAllowed(snapshotPath);

  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`File not found: ${filePath}`);
  }
  await fs.copyFile(filePath, snapshotPath);

  // Execute ops without rollback — we want to see what they actually did,
  // even on partial failure.
  const exec = await executeOperations(filePath, operations, {
    rollbackOnError: false,
    preCreatedSnapshot: snapshotPath,
    keepSnapshotOnSuccess: true, // we control deletion ourselves below
  });

  // Diff regardless of success/failure.
  const diffStr = await snapshotDiff(filePath, snapshotPath);
  const diff = JSON.parse(diffStr);

  // Optionally clean up the snapshot.
  let finalSnapshotPath: string | null = snapshotPath;
  if (!keepSnapshot) {
    await silentUnlink(snapshotPath);
    finalSnapshotPath = null;
  }

  if (!exec.ok) {
    return JSON.stringify(
      {
        success: false,
        failedAtStep: exec.failedAtStep,
        error: exec.error,
        operationsCompleted: exec.operationsCompleted,
        rolledBack: false,
        snapshotPath: finalSnapshotPath,
        diff,
      },
      null,
      2,
    );
  }

  return JSON.stringify(
    {
      success: true,
      operationsExecuted: exec.data.results.length,
      snapshotPath: finalSnapshotPath,
      diff,
    },
    null,
    2,
  );
}

// Help the index.ts wiring stay in sync.
export const TIER_C_TOOL_NAMES = [
  'excel_snapshot_create',
  'excel_snapshot_diff',
  'excel_snapshot_restore',
  'excel_transaction',
  'excel_diff_before_after',
] as const;

// Silence lint for unused `resolve` in case path manipulation grows.
void resolve;
