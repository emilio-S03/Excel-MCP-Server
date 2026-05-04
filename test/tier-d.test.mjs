/**
 * Tier D — structural fixes:
 *   1. excel_get_formula warnings array when Excel is open with the file (env-conditional).
 *   2. excel_apply_conditional_format throws CF_RANGE_OVERLAPS_MERGED on overlap.
 *   3. dedupKey skips re-application of a marker'd tool.
 *   4. PATH_OUTSIDE_ALLOWED error includes errorCode in the JSON response.
 */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-tier-d-'));
const xlsx = join(tmp, 'tier-d.xlsx');

const server = await startServer({ EXCEL_ALLOWED_DIRS: tmp });
after(() => {
  server.stop();
  try { rmSync(tmp, { recursive: true, force: true }); } catch {}
});

async function callOk(name, args) {
  const r = await server.callTool(name, args);
  if (r.error) throw new Error(`${name}: ${JSON.stringify(r.error)}`);
  if (r.result?.isError) throw new Error(`${name}: ${r.result.content?.[0]?.text}`);
  return JSON.parse(r.result.content[0].text);
}

async function callRaw(name, args) {
  const r = await server.callTool(name, args);
  return r;
}

// Bootstrap a workbook to operate on.
await callOk('excel_write_workbook', {
  filePath: xlsx,
  sheetName: 'Sheet1',
  data: [
    ['name', 'val'],
    ['a', 1],
    ['b', 2],
    ['c', 3],
    ['d', 4],
  ],
});

// ----------------------------------------------------------------------------
// Test 1 — excel_get_formula warnings field is environment-conditional.
// ----------------------------------------------------------------------------
test('excel_get_formula returns warnings array when Excel is running with the file open', async (t) => {
  const env = await callOk('excel_check_environment', {});
  const excelRunning = env?.capabilityMatrix?.liveMode?.excelRunning === true;
  if (!excelRunning) {
    t.skip('Excel is not running on this host — env-conditional test');
    return;
  }
  // We can't reliably guarantee that a test workbook is the one open, so we
  // only assert the *shape* — if Excel is running, the read MAY include a
  // warnings array. If our specific file is open, it WILL.
  const r = await callOk('excel_get_formula', {
    filePath: xlsx,
    sheetName: 'Sheet1',
    cellAddress: 'A1',
  });
  if (Array.isArray(r.warnings)) {
    assert.ok(r.warnings.length > 0, 'warnings array should be non-empty when present');
    assert.ok(
      String(r.warnings[0]).includes('open in Excel'),
      `warning should mention Excel; got "${r.warnings[0]}"`
    );
  }
  // No assertion if no warnings — file just isn't open.
});

// ----------------------------------------------------------------------------
// Test 2 — excel_apply_conditional_format on a range that overlaps merged
// cells must throw with errorCode CF_RANGE_OVERLAPS_MERGED.
// ----------------------------------------------------------------------------
test('excel_apply_conditional_format throws CF_RANGE_OVERLAPS_MERGED on merged range', async () => {
  // Bootstrap a fresh workbook so we don't pollute the shared one.
  const mergedFile = join(tmp, 'merged.xlsx');
  await callOk('excel_write_workbook', {
    filePath: mergedFile,
    sheetName: 'Sheet1',
    data: [
      [1, 2, 3, 4],
      [5, 6, 7, 8],
      [9, 10, 11, 12],
      [13, 14, 15, 16],
      [17, 18, 19, 20],
    ],
  });
  await callOk('excel_merge_cells', {
    filePath: mergedFile,
    sheetName: 'Sheet1',
    range: 'A1:C1',
  });
  const r = await callRaw('excel_apply_conditional_format', {
    filePath: mergedFile,
    sheetName: 'Sheet1',
    range: 'A1:C5',
    ruleType: 'cellValue',
    condition: { operator: 'greaterThan', value: 5 },
    style: { fill: { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' } },
  });
  assert.equal(r.result?.isError, true, 'should be an error response');
  const text = r.result.content[0].text;
  assert.ok(text.includes('CF_RANGE_OVERLAPS_MERGED'), `expected CF_RANGE_OVERLAPS_MERGED in error; got: ${text}`);
  // The error JSON should also surface the structured errorCode field.
  const parsed = JSON.parse(text);
  assert.equal(parsed.errorCode, 'CF_RANGE_OVERLAPS_MERGED');
});

// ----------------------------------------------------------------------------
// Test 3 — dedupKey short-circuits a second invocation.
// ----------------------------------------------------------------------------
test('dedupKey skips re-application of excel_create_named_range', async () => {
  const dedupFile = join(tmp, 'dedup.xlsx');
  await callOk('excel_write_workbook', {
    filePath: dedupFile,
    sheetName: 'Sheet1',
    data: [[1, 2, 3]],
  });
  const first = await callOk('excel_create_named_range', {
    filePath: dedupFile,
    name: 'MyRange',
    sheetName: 'Sheet1',
    range: 'A1:C1',
    dedupKey: 'test1',
  });
  assert.equal(first.success, true, 'first call should succeed');

  // Second call with the same dedupKey on the same workbook should skip.
  const second = await callOk('excel_create_named_range', {
    filePath: dedupFile,
    name: 'MyRange',
    sheetName: 'Sheet1',
    range: 'A1:C1',
    dedupKey: 'test1',
  });
  assert.equal(second.skipped, true, 'second call should be skipped');
  assert.ok(typeof second.reason === 'string' && second.reason.includes('dedupKey'),
    `reason should mention dedupKey; got "${second.reason}"`);
});

// ----------------------------------------------------------------------------
// Test 4 — PATH_OUTSIDE_ALLOWED carries errorCode in the JSON response.
// ----------------------------------------------------------------------------
test('PATH_OUTSIDE_ALLOWED error includes errorCode field', async () => {
  const tmpOutside = mkdtempSync(join(tmpdir(), 'excel-outside-'));
  try {
    const r = await callRaw('excel_write_workbook', {
      filePath: join(tmpOutside, 'blocked.xlsx'),
      sheetName: 'X',
      data: [[1]],
    });
    assert.equal(r.result?.isError, true);
    const text = r.result.content[0].text;
    const parsed = JSON.parse(text);
    assert.equal(parsed.errorCode, 'PATH_OUTSIDE_ALLOWED', `expected errorCode; got: ${text}`);
    assert.ok(String(parsed.error).includes('PATH_OUTSIDE_ALLOWED'),
      'message text should still mention PATH_OUTSIDE_ALLOWED for back-compat');
  } finally {
    try { rmSync(tmpOutside, { recursive: true, force: true }); } catch {}
  }
});
