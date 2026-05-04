/**
 * Idempotency markers (dedupKey).
 *
 * Each workbook can opt into duplicate-operation detection by passing a
 * `dedupKey` to certain mutating tools. The marker is recorded in a
 * `__excel_mcp_idempotency__` worksheet inside the workbook, hidden via
 * state: 'veryHidden' so users do not see it in the Sheets tab strip.
 *
 * Sheet schema:
 *   Row 1 (header): | dedupKey | toolName | timestamp | argsSummary |
 *   Row 2..N      : one row per recorded (dedupKey, toolName) pair.
 *
 * The same (dedupKey, toolName) pair recorded a second time means the
 * caller already applied this operation; the dispatcher should short-circuit
 * with skipped:true rather than re-invoking the underlying handler.
 */
import ExcelJS from 'exceljs';
import { ensureFilePathAllowed } from './helpers.js';
import { promises as fs } from 'fs';

export const IDEMPOTENCY_SHEET_NAME = '__excel_mcp_idempotency__';

export interface IdempotencyHit {
  skipped: boolean;
  previous?: { timestamp: string; argsSummary: string };
}

/**
 * Look up (or record) a dedupKey marker for a given tool. Loads the workbook
 * fresh, scans the very-hidden idempotency sheet, and either:
 *   - returns { skipped: true, previous } when the marker already exists, OR
 *   - appends a new row, saves, and returns { skipped: false }.
 *
 * Throws if the workbook can't be loaded or saved (path validation, missing
 * file, etc) — same failure modes as any other write tool.
 */
export async function checkAndRecord(
  filePath: string,
  toolName: string,
  dedupKey: string,
  argsSummary: string
): Promise<IdempotencyHit> {
  ensureFilePathAllowed(filePath);

  // The workbook must exist already — idempotency tracking is only meaningful
  // when the target workbook is real. If the file doesn't exist, we let the
  // underlying handler decide what to do (it'll typically auto-create).
  let fileExists = true;
  try {
    await fs.access(filePath);
  } catch {
    fileExists = false;
  }

  const workbook = new ExcelJS.Workbook();
  if (fileExists) {
    try {
      await workbook.xlsx.readFile(filePath);
    } catch (error) {
      // If we can't read the workbook we cannot track dedup either — bail
      // by returning skipped:false so the caller proceeds normally.
      return { skipped: false };
    }
  }

  let sheet = workbook.getWorksheet(IDEMPOTENCY_SHEET_NAME);
  if (!sheet) {
    sheet = workbook.addWorksheet(IDEMPOTENCY_SHEET_NAME, { state: 'veryHidden' });
    sheet.getRow(1).values = ['dedupKey', 'toolName', 'timestamp', 'argsSummary'];
    sheet.getRow(1).commit();
  } else {
    // Force veryHidden in case the sheet was made visible somehow.
    sheet.state = 'veryHidden';
    // Look up the marker.
    const lastRow = sheet.rowCount || 0;
    for (let r = 2; r <= lastRow; r++) {
      const row = sheet.getRow(r);
      const k = String(row.getCell(1).value ?? '');
      const t = String(row.getCell(2).value ?? '');
      if (k === dedupKey && t === toolName) {
        return {
          skipped: true,
          previous: {
            timestamp: String(row.getCell(3).value ?? ''),
            argsSummary: String(row.getCell(4).value ?? ''),
          },
        };
      }
    }
  }

  // Record a new marker.
  const newRowIdx = (sheet.rowCount || 1) + 1;
  const newRow = sheet.getRow(newRowIdx);
  newRow.getCell(1).value = dedupKey;
  newRow.getCell(2).value = toolName;
  newRow.getCell(3).value = new Date().toISOString();
  newRow.getCell(4).value = argsSummary;
  newRow.commit();

  await workbook.xlsx.writeFile(filePath);
  return { skipped: false };
}
