/**
 * Tier C — Snapshot tools (workbook-level point-in-time copies).
 *
 *   excel_snapshot_create   — copies a .xlsx to a sibling snapshot file
 *   excel_snapshot_diff     — compares all sheets between filePath and snapshot
 *   excel_snapshot_restore  — overwrites filePath with snapshot contents
 *
 * Snapshots ALWAYS live in the same directory as the source file so
 * ensureFilePathAllowed catches them automatically — they cannot be created
 * outside the sandbox.
 *
 * snapshot_diff reuses cell-comparison patterns from tier-b's compareSheets:
 * for each sheet that exists in both workbooks we walk every populated cell
 * (left ∪ right) and emit a diff entry whenever the value or formula differs.
 */
import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import { dirname, basename, join } from 'path';
import { constants as fsConstants } from 'fs';

import {
  ensureFilePathAllowed,
  loadWorkbook,
  cellValueToString,
} from './helpers.js';
import type { CellValue } from 'exceljs';

// ----------------------------------------------------------------------------
// snapshot_create
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
  const stem = base.replace(/\.(xlsx|xlsm)$/i, '');
  return join(dir, `${stem}.snapshot.${snapshotId}.xlsx`);
}

export async function snapshotCreate(
  filePath: string,
  snapshotId?: string,
): Promise<string> {
  ensureFilePathAllowed(filePath);

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
// snapshot_diff
// ----------------------------------------------------------------------------

function extractFormula(value: CellValue): string | null {
  if (value && typeof value === 'object' && 'formula' in value && (value as any).formula) {
    return String((value as any).formula);
  }
  if (value && typeof value === 'object' && 'sharedFormula' in value && (value as any).sharedFormula) {
    return String((value as any).sharedFormula);
  }
  return null;
}

interface CellInfo {
  value: any;       // friendly value (cached result for formulas; raw value otherwise)
  rawValue: CellValue;
  formula: string | null;
}

function collectCells(sheet: ExcelJS.Worksheet): Map<string, CellInfo> {
  const map = new Map<string, CellInfo>();
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const f = extractFormula(cell.value);
      let display: any = cell.value;
      if (f && cell.value && typeof cell.value === 'object' && 'result' in cell.value) {
        display = (cell.value as any).result;
      }
      map.set(cell.address, { value: display, rawValue: cell.value, formula: f });
    });
  });
  return map;
}

interface SheetDiffEntry {
  sheet: string;
  address: string;
  side: 'left-only' | 'right-only' | 'both-changed';
  leftValue?: any;
  rightValue?: any;
  leftFormula?: string;
  rightFormula?: string;
}

interface SheetDiffResult {
  added: number;
  removed: number;
  changed: number;
  diffs: SheetDiffEntry[];
}

function diffOneSheet(
  sheetName: string,
  leftSheet: ExcelJS.Worksheet,
  rightSheet: ExcelJS.Worksheet,
  remainingCap: number,
  includeValues: boolean,
  includeFormulas: boolean,
): SheetDiffResult {
  // "left" is the CURRENT (post-edit) workbook; "right" is the snapshot (the BEFORE state).
  // Per spec: left = current, snapshot = right (before)
  const leftMap = collectCells(leftSheet);
  const rightMap = collectCells(rightSheet);

  const allAddrs = new Set<string>([...leftMap.keys(), ...rightMap.keys()]);

  let added = 0;
  let removed = 0;
  let changed = 0;
  const diffs: SheetDiffEntry[] = [];

  for (const addr of allAddrs) {
    const l = leftMap.get(addr);
    const r = rightMap.get(addr);

    if (!l && r) {
      // Cell exists only on snapshot side (right) → REMOVED from current.
      removed++;
      if (diffs.length < remainingCap) {
        const entry: SheetDiffEntry = { sheet: sheetName, address: addr, side: 'right-only' };
        if (includeValues) entry.rightValue = r.value;
        if (includeFormulas && r.formula) entry.rightFormula = r.formula;
        diffs.push(entry);
      }
      continue;
    }

    if (l && !r) {
      // Only in current (left) → cell was ADDED since the snapshot
      added++;
      if (diffs.length < remainingCap) {
        const entry: SheetDiffEntry = { sheet: sheetName, address: addr, side: 'left-only' };
        if (includeValues) entry.leftValue = l.value;
        if (includeFormulas && l.formula) entry.leftFormula = l.formula;
        diffs.push(entry);
      }
      continue;
    }

    if (l && r) {
      const valueDiff = cellValueToString(l.rawValue) !== cellValueToString(r.rawValue);
      const formulaDiff = (l.formula ?? null) !== (r.formula ?? null);
      if (!valueDiff && !formulaDiff) continue;

      changed++;
      if (diffs.length < remainingCap) {
        const entry: SheetDiffEntry = { sheet: sheetName, address: addr, side: 'both-changed' };
        if (includeValues) {
          entry.leftValue = l.value;
          entry.rightValue = r.value;
        }
        if (includeFormulas) {
          if (l.formula) entry.leftFormula = l.formula;
          if (r.formula) entry.rightFormula = r.formula;
        }
        diffs.push(entry);
      }
    }
  }

  return { added, removed, changed, diffs };
}

export async function snapshotDiff(
  filePath: string,
  snapshotPath: string,
  options: {
    includeValues?: boolean;
    includeFormulas?: boolean;
    maxDifferences?: number;
  } = {},
): Promise<string> {
  ensureFilePathAllowed(filePath);
  ensureFilePathAllowed(snapshotPath);

  const includeValues = options.includeValues !== false;
  const includeFormulas = options.includeFormulas !== false;
  const maxDifferences = options.maxDifferences && options.maxDifferences > 0
    ? options.maxDifferences
    : 500;

  // current = "left", snapshot = "right" (before)
  const leftWb = await loadWorkbook(filePath);
  const rightWb = await loadWorkbook(snapshotPath);

  const leftSheets = new Map<string, ExcelJS.Worksheet>();
  leftWb.eachSheet((s) => leftSheets.set(s.name, s));
  const rightSheets = new Map<string, ExcelJS.Worksheet>();
  rightWb.eachSheet((s) => rightSheets.set(s.name, s));

  const allNames = new Set<string>([...leftSheets.keys(), ...rightSheets.keys()]);

  let sheetsAdded = 0;
  let sheetsRemoved = 0;
  let sheetsChanged = 0;
  let sheetsIdentical = 0;
  let totalCellChanges = 0;
  const perSheet: Array<any> = [];
  const differences: SheetDiffEntry[] = [];

  for (const name of allNames) {
    const leftSheet = leftSheets.get(name);
    const rightSheet = rightSheets.get(name);

    if (leftSheet && !rightSheet) {
      // Sheet exists in current but not in snapshot → ADDED since snapshot
      sheetsAdded++;
      const cells = collectCells(leftSheet);
      const addedCells = cells.size;
      totalCellChanges += addedCells;
      const remaining = Math.max(0, maxDifferences - differences.length);
      let pushed = 0;
      for (const [addr, info] of cells) {
        if (pushed >= remaining) break;
        const entry: SheetDiffEntry = { sheet: name, address: addr, side: 'left-only' };
        if (includeValues) entry.leftValue = info.value;
        if (includeFormulas && info.formula) entry.leftFormula = info.formula;
        differences.push(entry);
        pushed++;
      }
      perSheet.push({ sheetName: name, status: 'added', addedCells });
      continue;
    }

    if (rightSheet && !leftSheet) {
      // Sheet exists in snapshot but not in current → REMOVED since snapshot
      sheetsRemoved++;
      const cells = collectCells(rightSheet);
      const removedCells = cells.size;
      totalCellChanges += removedCells;
      const remaining = Math.max(0, maxDifferences - differences.length);
      let pushed = 0;
      for (const [addr, info] of cells) {
        if (pushed >= remaining) break;
        const entry: SheetDiffEntry = { sheet: name, address: addr, side: 'right-only' };
        if (includeValues) entry.rightValue = info.value;
        if (includeFormulas && info.formula) entry.rightFormula = info.formula;
        differences.push(entry);
        pushed++;
      }
      perSheet.push({ sheetName: name, status: 'removed', removedCells });
      continue;
    }

    // Both present.
    const remaining = Math.max(0, maxDifferences - differences.length);
    const sheetResult = diffOneSheet(
      name,
      leftSheet!,
      rightSheet!,
      remaining,
      includeValues,
      includeFormulas,
    );
    const cellChanges = sheetResult.added + sheetResult.removed + sheetResult.changed;
    totalCellChanges += cellChanges;
    if (cellChanges === 0) {
      sheetsIdentical++;
      perSheet.push({ sheetName: name, status: 'identical' });
    } else {
      sheetsChanged++;
      perSheet.push({
        sheetName: name,
        status: 'changed',
        addedCells: sheetResult.added,
        removedCells: sheetResult.removed,
        changedCells: sheetResult.changed,
      });
      for (const d of sheetResult.diffs) differences.push(d);
    }
  }

  const truncated = totalCellChanges > differences.length;

  return JSON.stringify(
    {
      summary: {
        sheetsAdded,
        sheetsRemoved,
        sheetsChanged,
        sheetsIdentical,
        totalCellChanges,
      },
      perSheet,
      differences,
      truncated,
      maxDifferences,
    },
    null,
    2,
  );
}

// ----------------------------------------------------------------------------
// snapshot_restore
// ----------------------------------------------------------------------------

function genBackupTimestamp(): string {
  return genTimestamp();
}

export async function snapshotRestore(
  filePath: string,
  snapshotPath: string,
  createBackup: boolean = true,
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
      const ts = genBackupTimestamp();
      const dir = dirname(filePath);
      const base = basename(filePath);
      const stem = base.replace(/\.(xlsx|xlsm)$/i, '');
      backupPath = join(dir, `${stem}.pre-restore-backup-${ts}.xlsx`);
      ensureFilePathAllowed(backupPath);
      await fs.copyFile(filePath, backupPath);
    } catch (err: any) {
      // If filePath doesn't exist, no backup needed; otherwise re-throw.
      if (err && err.code !== 'ENOENT') {
        // Only swallow ENOENT (file doesn't exist).
        if (!String(err.message ?? '').includes('ENOENT')) {
          // best-effort — continue
        }
      }
    }
  }

  // Try FICLONE first (fast copy-on-write on supported filesystems), fall back.
  try {
    await fs.copyFile(snapshotPath, filePath, fsConstants.COPYFILE_FICLONE);
  } catch {
    await fs.copyFile(snapshotPath, filePath);
  }

  const out: any = {
    success: true,
    restoredFrom: snapshotPath,
    restoredTo: filePath,
  };
  if (backupPath) out.preRestoreBackup = backupPath;
  return JSON.stringify(out, null, 2);
}

// Help the index.ts wiring stay in sync.
export const TIER_C_SNAPSHOT_TOOL_NAMES = [
  'excel_snapshot_create',
  'excel_snapshot_diff',
  'excel_snapshot_restore',
] as const;
