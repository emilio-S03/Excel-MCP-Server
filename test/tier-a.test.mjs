/**
 * Tier A bulk-operation tools — smoke tests (v3.3).
 *
 * Covers:
 *   1. excel_get_cell_styles_bulk
 *   2. excel_screenshot              (live-mode dispatch confirmation, no actual capture)
 *   3. excel_batch_write_formulas
 *   4. excel_get_data_validation_bulk — confirmed redundant; covered by excel_list_data_validations
 *   5. excel_create_named_range_bulk
 */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-tier-a-'));
const xlsxPath = join(tmp, 'tier-a.xlsx');
const csvIn = join(tmp, 'seed.csv');

writeFileSync(
  csvIn,
  'Name,Score,Status\nAlice,95,active\nBob,87,active\nCharlie,72,inactive\n',
  'utf8'
);

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
  return server.callTool(name, args);
}

// Bootstrap: import a CSV so we have a real sheet to work with.
await callOk('excel_csv_import', {
  csvPath: csvIn,
  targetXlsx: xlsxPath,
  sheetName: 'Data',
  hasHeader: true,
});

test('tools/list registers all four new tier-A tool names', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  for (const t of [
    'excel_get_cell_styles_bulk',
    'excel_batch_write_formulas',
    'excel_create_named_range_bulk',
    'excel_screenshot',
  ]) {
    assert.ok(names.has(t), `missing tier-A tool: ${t}`);
  }
});

test('excel_get_cell_styles_bulk returns one entry per non-empty cell', async () => {
  // Apply a known style to A1 so we can verify it round-trips.
  await callOk('excel_format_cell', {
    filePath: xlsxPath,
    sheetName: 'Data',
    cellAddress: 'A1',
    format: { font: { bold: true, color: 'FFFF0000' }, numFmt: '0.00' },
  });
  const r = await callOk('excel_get_cell_styles_bulk', {
    filePath: xlsxPath,
    sheetName: 'Data',
    range: 'A1:C4',
  });
  assert.equal(r.sheetName, 'Data');
  assert.equal(r.range, 'A1:C4');
  assert.ok(Array.isArray(r.cells), 'cells must be an array');
  // We have a 4x3 region of header+data, all should be populated.
  assert.ok(r.cells.length >= 4 && r.cells.length <= 12, `unexpected cell count: ${r.cells.length}`);
  const a1 = r.cells.find((c) => c.address === 'A1');
  assert.ok(a1, 'A1 must appear in the result');
  assert.equal(a1.hasValue, true);
  assert.ok(a1.font && a1.font.bold === true, 'A1 font.bold should round-trip');
});

test('excel_get_cell_styles_bulk includeEmpty=true expands the result', async () => {
  const sparse = await callOk('excel_get_cell_styles_bulk', {
    filePath: xlsxPath,
    sheetName: 'Data',
    range: 'F1:H3',
    includeEmpty: false,
  });
  const dense = await callOk('excel_get_cell_styles_bulk', {
    filePath: xlsxPath,
    sheetName: 'Data',
    range: 'F1:H3',
    includeEmpty: true,
  });
  assert.equal(dense.cells.length, 9, 'F1:H3 dense view has 3*3 cells');
  assert.ok(dense.cells.length > sparse.cells.length, 'includeEmpty should grow the result');
});

test('excel_batch_write_formulas: atomic apply on valid input', async () => {
  const r = await callOk('excel_batch_write_formulas', {
    filePath: xlsxPath,
    sheetName: 'Data',
    formulas: [
      { cell: 'D2', formula: 'B2*2' },
      { cell: 'D3', formula: 'B3*2' },
      { cell: 'D4', formula: '=SUM(B2:B4)' },
    ],
  });
  assert.equal(r.success, true);
  assert.equal(r.written, 3);
  assert.equal(r.formulas.length, 3);
  assert.ok(r.formulas[0].formula.startsWith('='));

  // Round-trip: each formula is now persisted.
  const f1 = await callOk('excel_get_formula', {
    filePath: xlsxPath, sheetName: 'Data', cellAddress: 'D2',
  });
  assert.ok(String(f1.formula || '').includes('B2*2'));
});

test('excel_batch_write_formulas: rejects invalid entry without writing anything', async () => {
  // First, capture D5 baseline (currently empty).
  const before = await callOk('excel_get_formula', {
    filePath: xlsxPath, sheetName: 'Data', cellAddress: 'D5',
  });

  // Submit a batch where one entry has unbalanced parens — must reject all.
  const r = await callRaw('excel_batch_write_formulas', {
    filePath: xlsxPath,
    sheetName: 'Data',
    formulas: [
      { cell: 'D5', formula: 'SUM(B2:B4' },   // unbalanced parens
    ],
  });
  assert.equal(r.result.isError, true, 'invalid batch must error');
  assert.ok(r.result.content[0].text.includes('unbalanced parentheses'));

  // D5 should still be empty/identical.
  const after = await callOk('excel_get_formula', {
    filePath: xlsxPath, sheetName: 'Data', cellAddress: 'D5',
  });
  assert.deepEqual(after.formula ?? null, before.formula ?? null);
});

test('excel_create_named_range_bulk: creates multiple in one call', async () => {
  const r = await callOk('excel_create_named_range_bulk', {
    filePath: xlsxPath,
    names: [
      { name: 'Scores', sheetName: 'Data', range: 'B2:B4' },
      { name: 'Names',  sheetName: 'Data', range: 'A2:A4' },
    ],
  });
  assert.equal(r.success, true);
  assert.equal(r.created, 2);

  // Verify via excel_list_named_ranges
  const list = await callOk('excel_list_named_ranges', { filePath: xlsxPath });
  const namesFound = list.namedRanges.map((n) => n.name);
  assert.ok(namesFound.includes('Scores'), 'Scores named range must persist');
  assert.ok(namesFound.includes('Names'), 'Names named range must persist');
});

test('excel_create_named_range_bulk: rejects when any sheet missing (transactional)', async () => {
  const r = await callRaw('excel_create_named_range_bulk', {
    filePath: xlsxPath,
    names: [
      { name: 'GoodOne', sheetName: 'Data',         range: 'A1:A2' },
      { name: 'BadOne',  sheetName: 'NopeSheetXYZ', range: 'A1:A2' },
    ],
  });
  assert.equal(r.result.isError, true);
  assert.ok(r.result.content[0].text.includes('NopeSheetXYZ'));

  // GoodOne must NOT have been created (transactional rollback).
  const list = await callOk('excel_list_named_ranges', { filePath: xlsxPath });
  const namesFound = list.namedRanges.map((n) => n.name);
  assert.ok(!namesFound.includes('GoodOne'), 'GoodOne must NOT exist after failed batch');
});

test('excel_screenshot: tool registered + dispatches to live mode', async () => {
  // We can't actually run a screenshot on a CI box without Excel + the file open.
  // The contract here: the tool exists, accepts the documented args, and reaches
  // the live-mode check (which may then return EXCEL_NOT_RUNNING — that's the
  // expected dispatch outcome and confirms the route is wired).
  const r = await callRaw('excel_screenshot', {
    filePath: xlsxPath,
    sheetName: 'Data',
    outputPath: join(tmp, 'shot.png'),
  });
  // Either the live-mode capture succeeds (Excel actually open — unlikely in CI)
  // or it returns EXCEL_NOT_RUNNING. Either way it must NOT be "Unknown tool".
  if (r.result.isError) {
    const txt = r.result.content[0].text;
    assert.ok(
      !/Unknown tool/.test(txt),
      `tool not wired into dispatcher: ${txt}`
    );
  }
});
