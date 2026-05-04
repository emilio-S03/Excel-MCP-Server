/**
 * Tier C — Transactional batch executor and convenience diff wrapper.
 *
 *   excel_transaction        — atomic batch executor (auto-rollback on failure)
 *   excel_diff_before_after  — convenience: snapshot + run ops + diff
 *
 * Both tools accept a list of {tool, args} operations and execute the underlying
 * tool handlers DIRECTLY (no MCP re-entry). Only file-mode write tools are
 * allowed — COM/AppleScript/PowerShell-backed tools are excluded because their
 * side effects on a running Excel process can't be rolled back by overwriting
 * the .xlsx file.
 */
import { promises as fs } from 'fs';
import { dirname, basename, join } from 'path';
import { randomBytes } from 'crypto';

import { ensureFilePathAllowed } from './helpers.js';

// Direct handler imports (file-mode write tools only — see SAFELIST docstring).
import { updateCell, writeRange, addRow, setFormula } from './write.js';
import { formatCell, mergeCells } from './format.js';
import { createSheet, deleteSheet, renameSheet } from './sheets.js';
import { deleteRows, deleteColumns, copyRange } from './operations.js';
import { insertRows, insertColumns, unmergeCells } from './advanced.js';
import { applyConditionalFormat } from './conditional.js';
import { createNamedRange, deleteNamedRange } from './named-ranges.js';
import { setDataValidation } from './validation.js';
import { sortRange, removeDuplicates, pasteSpecial } from './data-ops.js';
import { addHyperlink, removeHyperlink } from './hyperlinks.js';
import { csvImport, csvExport } from './csv.js';
import { findReplace } from './find-replace.js';
import { addImage } from './images.js';
import { setWorkbookProperties } from './inspections.js';
import { setSheetVisibility, hideRows, hideColumns } from './visibility.js';
import { setPageSetup } from './page-setup.js';
import { batchWriteFormulas, createNamedRangeBulk } from './bulk.js';

import { snapshotDiff } from './snapshot.js';

// ----------------------------------------------------------------------------
// Safelist
// ----------------------------------------------------------------------------

/**
 * Whitelist of tool names that excel_transaction (and excel_diff_before_after)
 * are allowed to invoke. Every entry here MUST be a file-mode (ExcelJS) write
 * tool — overwriting the .xlsx file is sufficient to roll back.
 *
 * Excluded from the safelist (and the reason):
 *   - excel_run_vba_macro / get_vba_code / set_vba_code (VBA execution against
 *     a live Excel process; side effects survive file rollback).
 *   - excel_create_modern_chart / create_combo_chart (live-mode COM only).
 *   - excel_capture_screenshot / excel_screenshot / excel_export_pdf (write to
 *     OS files outside the workbook; nothing to roll back, no value in
 *     transactional wrapping).
 *   - excel_trigger_recalculation / get_calculation_mode / set_calculation_mode
 *     (live-mode COM only).
 *   - excel_set_display_options (COM-backed in current implementation).
 *   - excel_add_shape (COM-backed in current implementation).
 *   - excel_set_auto_filter / clear_auto_filter (file-mode but side-effect-
 *     only; no value in batching them; can be added in a future bump).
 *   - excel_add_sparkline / remove_sparklines (XML-edit; safe to add later but
 *     deferred for v3.4 to keep surface area focused).
 *   - excel_set_sheet_protection / set_column_width / set_row_height /
 *     duplicate_sheet / add_comment / batch_format / set_calculation_mode
 *     (deliberately deferred — not on the spec's list).
 *
 * The map values are thin adapters that unpack the validated args and call
 * the underlying handler with positional arguments.
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
  excel_add_row: (a) => addRow(a.filePath, a.sheetName, a.data, a.createBackup),

  excel_create_sheet: (a) => createSheet(a.filePath, a.sheetName, a.createBackup),
  excel_delete_sheet: (a) => deleteSheet(a.filePath, a.sheetName, a.createBackup),
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

  excel_csv_import: (a) =>
    csvImport(a.csvPath, a.targetXlsx, {
      sheetName: a.sheetName,
      delimiter: a.delimiter,
      hasHeader: a.hasHeader,
      createBackup: a.createBackup,
    }),
  excel_csv_export: (a) =>
    csvExport(a.filePath, a.sheetName, a.csvPath, {
      range: a.range,
      delimiter: a.delimiter,
    }),

  excel_find_replace: (a) =>
    findReplace(a.filePath, a.pattern, a.replacement, {
      sheetName: a.sheetName,
      regex: a.regex,
      caseSensitive: a.caseSensitive,
      dryRun: a.dryRun,
      createBackup: a.createBackup,
    }),
  excel_add_image: (a) =>
    addImage(a.filePath, a.sheetName, a.imagePath, {
      cell: a.cell,
      range: a.range,
      widthPx: a.widthPx,
      heightPx: a.heightPx,
      createBackup: a.createBackup,
    }),

  excel_set_workbook_properties: (a) =>
    setWorkbookProperties(a.filePath, a.properties ?? a.props ?? {}, a.createBackup),

  excel_set_sheet_visibility: (a) =>
    setSheetVisibility(a.filePath, a.sheetName, a.state, a.createBackup),

  excel_hide_rows: (a) =>
    hideRows(a.filePath, a.sheetName, a.startRow, a.count, a.hidden, a.createBackup),
  excel_hide_columns: (a) =>
    hideColumns(
      a.filePath,
      a.sheetName,
      a.startColumn,
      a.count,
      a.hidden,
      a.createBackup,
    ),

  excel_set_page_setup: (a) =>
    setPageSetup(a.filePath, a.sheetName, a.config ?? a, a.createBackup),

  excel_batch_write_formulas: (a) =>
    batchWriteFormulas(a.filePath, a.sheetName, a.formulas, a.createBackup),
  excel_create_named_range_bulk: (a) =>
    createNamedRangeBulk(a.filePath, a.names, a.createBackup),
};

// ----------------------------------------------------------------------------
// Internals
// ----------------------------------------------------------------------------

interface OpSpec {
  tool: string;
  args: Record<string, any>;
}

function tempSnapshotPath(filePath: string): string {
  const dir = dirname(filePath);
  const base = basename(filePath);
  const stem = base.replace(/\.(xlsx|xlsm)$/i, '');
  const rand = randomBytes(6).toString('hex');
  // Hidden-ish prefix to make it obvious this is a transient file.
  return join(dir, `.${stem}.tx-${rand}.xlsx`);
}

async function silentUnlink(path: string): Promise<void> {
  try {
    await fs.unlink(path);
  } catch {
    /* ignore */
  }
}

class TransactionPlatformError extends Error {
  code = 'PLATFORM_UNSUPPORTED';
  constructor(message: string) {
    super(message);
    this.name = 'TransactionPlatformError';
  }
}

function rejectUnsafeTool(tool: string): never {
  throw new TransactionPlatformError(
    `Tool "${tool}" is not in the transaction safelist. ` +
      `Only file-mode write tools can be executed transactionally — COM/AppleScript/` +
      `PowerShell/live-mode tools are excluded because their side effects on a running ` +
      `Excel process can't be undone by file rollback. Safelist (${
        Object.keys(TRANSACTION_SAFELIST).length
      } tools): ${Object.keys(TRANSACTION_SAFELIST).sort().join(', ')}.`,
  );
}

function validateOpsOrThrow(operations: OpSpec[]): void {
  if (!Array.isArray(operations)) {
    throw new Error(`operations must be an array, got ${typeof operations}`);
  }
  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    if (!op || typeof op !== 'object' || typeof op.tool !== 'string') {
      throw new Error(
        `Operation ${i + 1} is malformed: expected { tool: string, args: object }, got ${JSON.stringify(op)}`,
      );
    }
    if (!(op.tool in TRANSACTION_SAFELIST)) {
      rejectUnsafeTool(op.tool);
    }
  }
}

function summarizeResult(raw: string): any {
  try {
    const parsed = JSON.parse(raw);
    // Trim large fields so transaction summaries stay terse.
    if (parsed && typeof parsed === 'object') {
      const summary: any = {};
      for (const k of [
        'success',
        'message',
        'rowsImported',
        'rowsExported',
        'matchCount',
        'mode',
        'written',
        'created',
        'sheetName',
        'state',
        'action',
      ]) {
        if (k in parsed) summary[k] = parsed[k];
      }
      return Object.keys(summary).length > 0 ? summary : { ok: true };
    }
  } catch {
    /* fall through */
  }
  return { ok: true };
}

// ----------------------------------------------------------------------------
// excel_transaction
// ----------------------------------------------------------------------------

export async function transaction(
  filePath: string,
  operations: OpSpec[],
  _createBackup: boolean = false,
): Promise<string> {
  ensureFilePathAllowed(filePath);
  // _createBackup is accepted for symmetry with other write tools; the snapshot
  // already serves as the rollback target so the param has no additional effect.
  void _createBackup;

  // Validate ALL ops up-front before we touch the disk.
  validateOpsOrThrow(operations);

  // Confirm source exists.
  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`File not found: ${filePath}`);
  }

  const snapshotPath = tempSnapshotPath(filePath);
  ensureFilePathAllowed(snapshotPath);
  await fs.copyFile(filePath, snapshotPath);

  const results: Array<{ step: number; tool: string; summary: any }> = [];

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    const exec = TRANSACTION_SAFELIST[op.tool];
    // Force-set filePath so callers can't smuggle a different target into a transactional call.
    const args = { ...(op.args ?? {}), filePath };
    try {
      const raw = await exec(args);
      results.push({ step: i + 1, tool: op.tool, summary: summarizeResult(raw) });
    } catch (err: any) {
      const errorMessage = err instanceof Error ? err.message : String(err);
      const errorCode = err && typeof err === 'object' && 'code' in err ? (err as any).code : undefined;

      // Restore from snapshot.
      try {
        await fs.copyFile(snapshotPath, filePath);
      } catch (restoreErr: any) {
        // Catastrophic: rollback failed. Preserve the snapshot for manual recovery.
        return JSON.stringify(
          {
            success: false,
            failedAtStep: i + 1,
            error: `${errorMessage} (rollback FAILED: ${
              restoreErr?.message ?? restoreErr
            }; snapshot preserved at ${snapshotPath})`,
            errorCode,
            operationsCompleted: i,
            rolledBack: false,
            snapshotPath,
          },
          null,
          2,
        );
      }
      await silentUnlink(snapshotPath);
      return JSON.stringify(
        {
          success: false,
          failedAtStep: i + 1,
          error: errorMessage,
          ...(errorCode ? { errorCode } : {}),
          operationsCompleted: i,
          rolledBack: true,
        },
        null,
        2,
      );
    }
  }

  // All ops succeeded — clean up the snapshot.
  await silentUnlink(snapshotPath);
  return JSON.stringify(
    {
      success: true,
      operationsExecuted: results.length,
      results,
    },
    null,
    2,
  );
}

// ----------------------------------------------------------------------------
// excel_diff_before_after
// ----------------------------------------------------------------------------

export async function diffBeforeAfter(
  filePath: string,
  operations: OpSpec[],
  keepSnapshot: boolean = false,
): Promise<string> {
  ensureFilePathAllowed(filePath);

  validateOpsOrThrow(operations);

  try {
    await fs.access(filePath);
  } catch {
    throw new Error(`File not found: ${filePath}`);
  }

  const snapshotPath = tempSnapshotPath(filePath);
  ensureFilePathAllowed(snapshotPath);
  await fs.copyFile(filePath, snapshotPath);

  let operationsExecuted = 0;
  let firstError: { failedAtStep: number; error: string; errorCode?: string } | null = null;

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    const exec = TRANSACTION_SAFELIST[op.tool];
    const args = { ...(op.args ?? {}), filePath };
    try {
      await exec(args);
      operationsExecuted++;
    } catch (err: any) {
      const errorMessage = err instanceof Error ? err.message : String(err);
      const errorCode = err && typeof err === 'object' && 'code' in err ? (err as any).code : undefined;
      firstError = {
        failedAtStep: i + 1,
        error: errorMessage,
        ...(errorCode ? { errorCode } : {}),
      };
      // diff_before_after does NOT roll back — we want to see what the partial
      // run actually changed.
      break;
    }
  }

  // Diff regardless of success/failure.
  const diffStr = await snapshotDiff(filePath, snapshotPath);
  const diff = JSON.parse(diffStr);

  let finalSnapshotPath: string | null = snapshotPath;
  if (!keepSnapshot) {
    await silentUnlink(snapshotPath);
    finalSnapshotPath = null;
  }

  const out: any = {
    success: firstError === null,
    operationsExecuted,
    diff,
  };
  if (finalSnapshotPath) out.snapshotPath = finalSnapshotPath;
  if (firstError) {
    out.failedAtStep = firstError.failedAtStep;
    out.error = firstError.error;
    if ('errorCode' in firstError) out.errorCode = (firstError as any).errorCode;
  }
  return JSON.stringify(out, null, 2);
}

export const TIER_C_TRANSACTION_TOOL_NAMES = [
  'excel_transaction',
  'excel_diff_before_after',
] as const;
