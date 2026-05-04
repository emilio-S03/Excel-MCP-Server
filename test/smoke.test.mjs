import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, readFileSync, existsSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-smoke-'));
const csvIn = join(tmp, 'in.csv');
const csvOut = join(tmp, 'out.csv');
const xlsxPath = join(tmp, 'sheet.xlsx');

writeFileSync(
  csvIn,
  'Name,Age,Score,Status\nAlice,30,95.5,active\nBob,25,87.0,active\nCharlie,40,72.3,inactive\nDana,35,88.0,active\nEve,28,91.0,inactive\n',
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

test('tools/list exposes 104 tools including v3 + v3.1 + modern-charts + audit + tier-A + tier-B additions', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  assert.equal(tools.length, 104);
  for (const t of [
    'excel_check_environment',
    'excel_add_image',
    'excel_csv_import',
    'excel_csv_export',
    'excel_find_replace',
  ]) {
    assert.ok(names.has(t), `missing tool: ${t}`);
  }
});

test('excel_check_environment returns capability matrix', async () => {
  const r = await callOk('excel_check_environment', {});
  assert.ok(['win32', 'darwin', 'linux'].includes(r.platform));
  assert.equal(r.serverVersion, '3.3.0');
  assert.equal(r.capabilityMatrix.fileMode.available, true);
  assert.ok(Array.isArray(r.config.allowedDirectories));
  assert.equal(r.config.allowedDirectoriesIsDefault, false, 'sandbox is overridden by test env');
});

test('excel_csv_import creates workbook from CSV', async () => {
  const r = await callOk('excel_csv_import', {
    csvPath: csvIn,
    targetXlsx: xlsxPath,
    sheetName: 'People',
    hasHeader: true,
  });
  assert.equal(r.success, true);
  assert.equal(r.rowsImported, 6);
  assert.ok(existsSync(xlsxPath));
});

test('excel_read_sheet pagination returns first page + nextOffset', async () => {
  const r = await callOk('excel_read_sheet', {
    filePath: xlsxPath,
    sheetName: 'People',
    limit: 2,
  });
  assert.equal(r.rowCount, 2);
  assert.equal(r.totalRows, 6);
  assert.equal(r.hasMore, true);
  assert.equal(r.nextOffset, 2);
});

test('excel_read_sheet offset jumps in', async () => {
  const r = await callOk('excel_read_sheet', {
    filePath: xlsxPath,
    sheetName: 'People',
    offset: 4,
  });
  assert.equal(r.rowCount, 2);
  assert.equal(r.hasMore, false);
  assert.equal(r.nextOffset, null);
});

test('excel_find_replace dryRun finds substring matches without modifying', async () => {
  const r = await callOk('excel_find_replace', {
    filePath: xlsxPath,
    pattern: 'active',
    replacement: 'ENABLED',
    caseSensitive: true,
    dryRun: true,
  });
  assert.equal(r.matchCount, 5); // "active" matches inside "inactive" too
  assert.equal(r.mode, 'dryRun');

  const sheet = await callOk('excel_read_sheet', { filePath: xlsxPath, sheetName: 'People' });
  assert.ok(!JSON.stringify(sheet.rows).includes('ENABLED'));
});

test('excel_find_replace applies changes', async () => {
  const r = await callOk('excel_find_replace', {
    filePath: xlsxPath,
    pattern: 'inactive',
    replacement: 'DISABLED',
  });
  assert.equal(r.matchCount, 2);
  assert.equal(r.mode, 'applied');

  const sheet = await callOk('excel_read_sheet', { filePath: xlsxPath, sheetName: 'People' });
  const flat = JSON.stringify(sheet.rows);
  assert.ok(flat.includes('DISABLED'));
  assert.ok(!flat.includes('inactive'));
});

test('excel_csv_export round-trips', async () => {
  const r = await callOk('excel_csv_export', {
    filePath: xlsxPath,
    sheetName: 'People',
    csvPath: csvOut,
  });
  assert.equal(r.success, true);
  const out = readFileSync(csvOut, 'utf8');
  assert.ok(out.includes('Alice'));
  assert.ok(out.includes('DISABLED'));
});

test('excel_find_replace regex with backreference', async () => {
  const r = await callOk('excel_find_replace', {
    filePath: xlsxPath,
    pattern: 'A(\\w+)',
    replacement: 'X$1',
    regex: true,
    caseSensitive: true,
    dryRun: true,
  });
  assert.ok(r.matchCount >= 1);
  assert.ok(r.matches.some((m) => m.after.startsWith('X')));
});
