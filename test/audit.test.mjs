import { test, after, before } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-audit-'));
const xlsxPath = join(tmp, 'audit.xlsx');

// Bootstrap a workbook with the fixtures the task spec calls out:
// - normal SUM formula in B1 referencing A1:A5
// - #DIV/0! error in C1 (formula =1/0)
// - circular reference in D1 (formula referencing D1 itself)
// - a named range over A1:A5
before(async () => {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('Data');

  sheet.getCell('A1').value = 1;
  sheet.getCell('A2').value = 2;
  sheet.getCell('A3').value = 3;
  sheet.getCell('A4').value = 4;
  sheet.getCell('A5').value = 5;

  // Normal formula
  sheet.getCell('B1').value = { formula: 'SUM(A1:A5)', result: 15 };

  // #DIV/0! error — store both formula and the error result so detectors fire.
  sheet.getCell('C1').value = { formula: '1/0', result: { error: '#DIV/0!' } };

  // Circular reference — D1 references itself.
  sheet.getCell('D1').value = { formula: 'D1+1', result: 0 };

  // Named range over A1:A5
  wb.definedNames.add(`'Data'!$A$1:$A$5`, 'MyNumbers');

  await wb.xlsx.writeFile(xlsxPath);
});

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

test('tools/list reports 110 tools including formula audit + tier-A + tier-B + tier-C additions', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  assert.equal(tools.length, 110);
  for (const t of [
    'excel_find_formula_errors',
    'excel_find_circular_references',
    'excel_workbook_stats',
    'excel_list_formulas',
    'excel_trace_precedents',
  ]) {
    assert.ok(names.has(t), `missing tool: ${t}`);
  }
});

test('excel_find_formula_errors returns the #DIV/0! cell', async () => {
  const r = await callOk('excel_find_formula_errors', { filePath: xlsxPath });
  assert.ok(r.totalErrors >= 1, `expected at least one error, got ${r.totalErrors}`);
  const div0 = r.errors.find((e) => e.cell === 'C1');
  assert.ok(div0, 'expected an entry for C1');
  assert.match(div0.errorType, /#DIV\/0!?/);
  assert.equal(div0.sheet, 'Data');
});

test('excel_find_formula_errors with sheetName scopes to that sheet', async () => {
  const r = await callOk('excel_find_formula_errors', { filePath: xlsxPath, sheetName: 'Data' });
  assert.ok(r.totalErrors >= 1);
  for (const err of r.errors) assert.equal(err.sheet, 'Data');
});

test('excel_find_circular_references returns D1 self-reference', async () => {
  const r = await callOk('excel_find_circular_references', { filePath: xlsxPath });
  assert.ok(r.totalCircular >= 1, `expected at least one circular ref, got ${r.totalCircular}`);
  const d1 = r.references.find((c) => c.cell === 'D1');
  assert.ok(d1, 'expected a circular reference entry for D1');
  assert.equal(d1.formula, 'D1+1');
  assert.ok(d1.referencedCells.includes('Data!D1') || d1.referencedCells.includes('D1'));
});

test('excel_workbook_stats returns plausible counts', async () => {
  const r = await callOk('excel_workbook_stats', { filePath: xlsxPath });
  assert.equal(r.totalSheets, 1);
  // 5 numbers in A1:A5 + B1 formula + C1 formula + D1 formula = 8 cells
  assert.ok(r.totalCells >= 8, `expected >= 8 cells, got ${r.totalCells}`);
  assert.ok(r.formulaCells >= 3, `expected >= 3 formula cells, got ${r.formulaCells}`);
  assert.equal(r.namedRanges, 1);
  assert.ok(r.fileSizeBytes > 0);
  assert.equal(r.sheetStats.length, 1);
  assert.equal(r.sheetStats[0].sheet, 'Data');
  assert.ok(r.sheetStats[0].cellsUsed >= 8);
  assert.ok(r.sheetStats[0].formulasUsed >= 3);
});

test('excel_list_formulas includes B1 SUM(A1:A5)', async () => {
  const r = await callOk('excel_list_formulas', { filePath: xlsxPath, sheetName: 'Data' });
  assert.ok(r.totalFormulas >= 3, `expected >= 3 formulas, got ${r.totalFormulas}`);
  const b1 = r.formulas.find((f) => f.cell === 'B1');
  assert.ok(b1, 'expected B1 in formula list');
  assert.equal(b1.formula, 'SUM(A1:A5)');
});

test('excel_list_formulas respects maxResults cap', async () => {
  const r = await callOk('excel_list_formulas', {
    filePath: xlsxPath,
    sheetName: 'Data',
    maxResults: 1,
  });
  assert.equal(r.totalFormulas, 1);
  assert.equal(r.truncated, true);
});

test('excel_trace_precedents on B1 returns A1:A5 references', async () => {
  const r = await callOk('excel_trace_precedents', {
    filePath: xlsxPath,
    sheetName: 'Data',
    cellAddress: 'B1',
  });
  assert.equal(r.cell, 'B1');
  assert.equal(r.formula, 'SUM(A1:A5)');
  assert.equal(r.depth, 1);
  // Should expand A1:A5 into 5 precedent entries.
  const cellAddrs = new Set(r.directPrecedents.map((p) => p.cell));
  for (const addr of ['A1', 'A2', 'A3', 'A4', 'A5']) {
    assert.ok(cellAddrs.has(addr), `expected precedent ${addr}`);
  }
  // Each precedent should resolve to its numeric value.
  const a1 = r.directPrecedents.find((p) => p.cell === 'A1');
  assert.equal(a1.value, 1);
});

test('excel_trace_precedents on a non-formula cell returns empty list', async () => {
  const r = await callOk('excel_trace_precedents', {
    filePath: xlsxPath,
    sheetName: 'Data',
    cellAddress: 'A1',
  });
  assert.equal(r.formula, null);
  assert.deepEqual(r.directPrecedents, []);
});
