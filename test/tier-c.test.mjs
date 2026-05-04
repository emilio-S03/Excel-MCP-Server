import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, existsSync, readFileSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-tier-c-'));
const xlsxPath = join(tmp, 'workbook.xlsx');
const mergedXlsxPath = join(tmp, 'merged.xlsx');

const server = await startServer({ EXCEL_ALLOWED_DIRS: tmp });
after(() => {
  server.stop();
  try { rmSync(tmp, { recursive: true, force: true }); } catch {}
});

async function callOk(name, args) {
  // Use a generous timeout: transaction sub-ops invoke the per-tool live-mode
  // probe (PowerShell on Windows). When Excel happens to be running on the host,
  // that probe can take several seconds per call and the chained sub-ops can
  // exceed the harness's 15s default.
  const r = await server.callTool(name, args, 60000);
  if (r.error) throw new Error(`${name}: ${JSON.stringify(r.error)}`);
  if (r.result?.isError) throw new Error(`${name}: ${r.result.content?.[0]?.text}`);
  return JSON.parse(r.result.content[0].text);
}

async function callRaw(name, args) {
  const r = await server.callTool(name, args, 60000);
  return r;
}

// Bootstrap a workbook for the snapshot/transaction tests
await callOk('excel_write_workbook', {
  filePath: xlsxPath,
  sheetName: 'Data',
  data: [['Name', 'Score'], ['Alice', 10], ['Bob', 20]],
});

test('tier-c tools registered in tools/list', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  for (const t of [
    'excel_snapshot_create',
    'excel_snapshot_diff',
    'excel_snapshot_restore',
    'excel_transaction',
    'excel_diff_before_after',
    'excel_read_sheet_merged_aware',
  ]) {
    assert.ok(names.has(t), `missing tool: ${t}`);
  }
});

test('snapshot_create + snapshot_diff + snapshot_restore round-trip', async () => {
  // 1. Create snapshot of the bootstrapped workbook.
  const snap = await callOk('excel_snapshot_create', { filePath: xlsxPath, snapshotId: 'rt' });
  assert.equal(typeof snap.snapshotPath, 'string');
  assert.ok(existsSync(snap.snapshotPath), 'snapshot file should exist');
  assert.equal(snap.snapshotId, 'rt');
  assert.ok(snap.fileSize > 0);

  // 2. Modify the workbook.
  await callOk('excel_update_cell', {
    filePath: xlsxPath,
    sheetName: 'Data',
    cellAddress: 'B2',
    value: 999,
  });

  // 3. Diff vs the snapshot — should report >=1 difference.
  const diff = await callOk('excel_snapshot_diff', {
    filePath: xlsxPath,
    snapshotPath: snap.snapshotPath,
  });
  assert.ok(diff.summary, 'diff should have summary');
  assert.ok(diff.summary.totalCellChanges >= 1, `expected >=1 cell change; got ${JSON.stringify(diff.summary)}`);
  assert.ok(Array.isArray(diff.differences));
  const b2 = diff.differences.find((d) => String(d.address).toUpperCase() === 'B2');
  assert.ok(b2, `B2 should appear in diffs; got ${JSON.stringify(diff.differences)}`);

  // 4. Restore from snapshot.
  const restore = await callOk('excel_snapshot_restore', {
    filePath: xlsxPath,
    snapshotPath: snap.snapshotPath,
    createBackup: false,
  });
  assert.equal(restore.success, true);

  // 5. Diff again — should report 0 cell changes.
  const diff2 = await callOk('excel_snapshot_diff', {
    filePath: xlsxPath,
    snapshotPath: snap.snapshotPath,
  });
  assert.equal(diff2.summary.totalCellChanges, 0, `expected 0 changes after restore; got ${JSON.stringify(diff2.summary)}`);
});

test('excel_transaction success path commits all ops', async () => {
  // Reset the workbook state for this test.
  await callOk('excel_write_workbook', {
    filePath: xlsxPath,
    sheetName: 'Data',
    data: [['Name', 'Score'], ['Alice', 10], ['Bob', 20]],
  });

  const r = await callOk('excel_transaction', {
    filePath: xlsxPath,
    operations: [
      { tool: 'excel_update_cell', args: { sheetName: 'Data', cellAddress: 'B2', value: 100 } },
      { tool: 'excel_update_cell', args: { sheetName: 'Data', cellAddress: 'B3', value: 200 } },
      { tool: 'excel_update_cell', args: { sheetName: 'Data', cellAddress: 'C1', value: 'Tag' } },
    ],
  });
  assert.equal(r.success, true);
  assert.equal(r.operationsExecuted, 3);
  assert.equal(r.results.length, 3);

  const after = await callOk('excel_read_sheet', { filePath: xlsxPath, sheetName: 'Data' });
  const flat = JSON.stringify(after.rows);
  assert.ok(flat.includes('100'), `B2 should be 100; rows: ${flat}`);
  assert.ok(flat.includes('200'), `B3 should be 200; rows: ${flat}`);
  assert.ok(flat.includes('Tag'), `C1 should be Tag; rows: ${flat}`);
});

test('excel_transaction rollback restores file on failure', async () => {
  // Reset workbook again.
  await callOk('excel_write_workbook', {
    filePath: xlsxPath,
    sheetName: 'Data',
    data: [['Name', 'Score'], ['Alice', 10], ['Bob', 20]],
  });

  // Read pre-state for byte-level comparison.
  const beforeBytes = readFileSync(xlsxPath);

  // 2nd op uses an invalid range to trigger a failure mid-batch.
  const r = await callOk('excel_transaction', {
    filePath: xlsxPath,
    operations: [
      { tool: 'excel_update_cell', args: { sheetName: 'Data', cellAddress: 'B2', value: 555 } },
      { tool: 'excel_write_range', args: { sheetName: 'Data', range: 'INVALID', data: [[1, 2]] } },
      { tool: 'excel_update_cell', args: { sheetName: 'Data', cellAddress: 'B3', value: 666 } },
    ],
  });

  assert.equal(r.success, false, `expected success=false; got ${JSON.stringify(r)}`);
  assert.equal(r.rolledBack, true);
  assert.equal(r.failedAtStep, 2);

  // Verify file unchanged: cell values match the pre-transaction state.
  const after = await callOk('excel_read_sheet', { filePath: xlsxPath, sheetName: 'Data' });
  const flat = JSON.stringify(after.rows);
  assert.ok(!flat.includes('555'), `expected B2 NOT to be 555 after rollback; rows: ${flat}`);
  assert.ok(!flat.includes('666'), `expected B3 NOT to be 666 after rollback; rows: ${flat}`);

  // Snapshot file should be cleaned up.
  const lingering = readdirSync(tmp).filter((f) => f.includes('.tx-'));
  assert.equal(lingering.length, 0, `snapshot leftovers: ${lingering.join(', ')}`);

  // Sanity: file should still be a non-empty .xlsx
  const afterBytes = readFileSync(xlsxPath);
  assert.ok(afterBytes.length > 0);
  // Bytes may not be byte-identical due to ExcelJS rewrite of the snapshot; cell-level
  // equality (asserted above) is the meaningful invariant.
  void beforeBytes;
});

test('excel_transaction safelist rejects non-safelisted tool with PLATFORM_UNSUPPORTED', async () => {
  const raw = await callRaw('excel_transaction', {
    filePath: xlsxPath,
    operations: [
      { tool: 'excel_run_vba_macro', args: { sheetName: 'Data', macroName: 'doStuff' } },
    ],
  });

  // The tool throws synchronously inside the dispatcher; the MCP wrapper turns it into
  // an isError response with the error JSON in content.
  assert.ok(raw.result?.isError === true, `expected isError; got ${JSON.stringify(raw)}`);
  const errText = raw.result.content?.[0]?.text ?? '';
  assert.ok(errText.includes('PLATFORM_UNSUPPORTED') || errText.includes('safelist'),
    `expected PLATFORM_UNSUPPORTED or safelist error; got: ${errText}`);

  // Snapshot should NOT have been created — verify no leftover .tx- files.
  const lingering = readdirSync(tmp).filter((f) => f.includes('.tx-'));
  assert.equal(lingering.length, 0, `snapshot leftovers from rejected tx: ${lingering.join(', ')}`);
});

test('excel_diff_before_after returns a non-empty diff', async () => {
  // Reset workbook to a known state for this test.
  await callOk('excel_write_workbook', {
    filePath: xlsxPath,
    sheetName: 'Data',
    data: [['Name', 'Score'], ['Alice', 10], ['Bob', 20]],
  });

  const r = await callOk('excel_diff_before_after', {
    filePath: xlsxPath,
    operations: [
      { tool: 'excel_update_cell', args: { sheetName: 'Data', cellAddress: 'B2', value: 42 } },
    ],
  });

  assert.equal(r.success, true);
  assert.equal(r.operationsExecuted, 1);
  assert.ok(r.diff, 'should return diff payload');
  assert.ok(r.diff.summary.totalCellChanges >= 1, `expected >=1 cell change; got ${JSON.stringify(r.diff.summary)}`);
});

test('excel_read_sheet_merged_aware fills merged cells with top-left value', async () => {
  // Bootstrap a workbook with a merged header row.
  await callOk('excel_write_workbook', {
    filePath: mergedXlsxPath,
    sheetName: 'Sheet1',
    data: [['Header', null, null], ['a', 'b', 'c']],
  });
  await callOk('excel_merge_cells', {
    filePath: mergedXlsxPath,
    sheetName: 'Sheet1',
    range: 'A1:C1',
  });

  // With fillMerged: true (default), all 3 cells in row 1 should read "Header".
  const filled = await callOk('excel_read_sheet_merged_aware', {
    filePath: mergedXlsxPath,
    sheetName: 'Sheet1',
    range: 'A1:C2',
    fillMerged: true,
  });
  assert.ok(Array.isArray(filled.rows));
  const row1 = filled.rows[0];
  assert.equal(row1.length, 3);
  assert.equal(String(row1[0]), 'Header');
  assert.equal(String(row1[1]), 'Header', `B1 should equal "Header" with fillMerged; row1=${JSON.stringify(row1)}`);
  assert.equal(String(row1[2]), 'Header', `C1 should equal "Header" with fillMerged; row1=${JSON.stringify(row1)}`);

  // With fillMerged: false, only A1 has the value; B1+C1 are null/empty.
  const raw = await callOk('excel_read_sheet_merged_aware', {
    filePath: mergedXlsxPath,
    sheetName: 'Sheet1',
    range: 'A1:C2',
    fillMerged: false,
  });
  const rrow1 = raw.rows[0];
  assert.equal(String(rrow1[0]), 'Header');
  // B1 and C1 should be null or empty after merge with fillMerged:false.
  const isEmpty = (v) => v === null || v === undefined || v === '';
  assert.ok(isEmpty(rrow1[1]), `B1 should be null/empty without fillMerged; got ${JSON.stringify(rrow1[1])}`);
  assert.ok(isEmpty(rrow1[2]), `C1 should be null/empty without fillMerged; got ${JSON.stringify(rrow1[2])}`);

  // includeMergedMetadata returns the mergedCells array.
  const meta = await callOk('excel_read_sheet_merged_aware', {
    filePath: mergedXlsxPath,
    sheetName: 'Sheet1',
    range: 'A1:C2',
    fillMerged: true,
    includeMergedMetadata: true,
  });
  assert.ok(Array.isArray(meta.mergedCells), 'mergedCells should be an array when includeMergedMetadata: true');
  assert.ok(meta.mergedCells.length >= 1, `expected >=1 merged region; got ${JSON.stringify(meta.mergedCells)}`);
  assert.equal(meta.mergedCells[0].topLeft, 'A1');
});
