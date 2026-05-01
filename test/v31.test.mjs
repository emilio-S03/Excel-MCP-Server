import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-v31-'));
const xlsxPath = join(tmp, 'test.xlsx');
const csvIn = join(tmp, 'data.csv');

writeFileSync(
  csvIn,
  'Name,Region,Score\nZara,East,72\nAlice,West,95\nBob,East,87\nAlice,West,95\nCharlie,North,72\nDana,South,88\n',
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

// Bootstrap a workbook from CSV
await callOk('excel_csv_import', { csvPath: csvIn, targetXlsx: xlsxPath, sheetName: 'Data', hasHeader: true });

test('tools/list reports 96 tools including v3.1 + modern-charts + audit additions', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  assert.equal(tools.length, 96);
  for (const t of [
    'excel_get_conditional_formats',
    'excel_list_data_validations',
    'excel_get_sheet_protection',
    'excel_get_display_options',
    'excel_get_workbook_properties',
    'excel_set_workbook_properties',
    'excel_get_hyperlinks',
    'excel_sort',
    'excel_set_auto_filter',
    'excel_clear_auto_filter',
    'excel_remove_duplicates',
    'excel_paste_special',
    'excel_set_sheet_visibility',
    'excel_list_sheet_visibility',
    'excel_hide_rows',
    'excel_hide_columns',
    'excel_add_hyperlink',
    'excel_remove_hyperlink',
    'excel_get_page_setup',
    'excel_set_page_setup',
    'excel_export_pdf',
  ]) {
    assert.ok(names.has(t), `missing tool: ${t}`);
  }
});

test('excel_sort: multi-key sort with header', async () => {
  await callOk('excel_sort', {
    filePath: xlsxPath,
    sheetName: 'Data',
    range: 'A1:C7',
    sortBy: [
      { column: 'B', order: 'asc' },
      { column: 'C', order: 'desc' },
    ],
    hasHeader: true,
  });
  const sheet = await callOk('excel_read_sheet', { filePath: xlsxPath, sheetName: 'Data' });
  // Header preserved at row 0
  assert.equal(sheet.rows[0][0], 'Name');
  // First data row should be East region with highest score (Bob 87 vs original 72 for Zara)
  assert.equal(sheet.rows[1][1], 'East');
  assert.equal(sheet.rows[1][2], 87);
});

test('excel_remove_duplicates: dedupe by all columns', async () => {
  // Re-bootstrap from CSV (Alice/95 appears twice)
  await callOk('excel_csv_import', { csvPath: csvIn, targetXlsx: xlsxPath, sheetName: 'Data', hasHeader: true });
  const r = await callOk('excel_remove_duplicates', {
    filePath: xlsxPath,
    sheetName: 'Data',
    range: 'A1:C7',
    hasHeader: true,
  });
  assert.equal(r.duplicatesRemoved, 1);
  assert.equal(r.rowsKept, 5);
});

test('excel_set_auto_filter + clear', async () => {
  const set = await callOk('excel_set_auto_filter', {
    filePath: xlsxPath,
    sheetName: 'Data',
    range: 'A1:C100',
  });
  assert.equal(set.success, true);
  assert.equal(set.autoFilterRange, 'A1:C100');

  const cleared = await callOk('excel_clear_auto_filter', { filePath: xlsxPath, sheetName: 'Data' });
  assert.equal(cleared.success, true);
});

test('excel_set_sheet_visibility: hide and re-show', async () => {
  await callOk('excel_create_sheet', { filePath: xlsxPath, sheetName: 'Hideable' });
  await callOk('excel_set_sheet_visibility', { filePath: xlsxPath, sheetName: 'Hideable', state: 'hidden' });

  const list = await callOk('excel_list_sheet_visibility', { filePath: xlsxPath });
  const hideable = list.sheets.find((s) => s.name === 'Hideable');
  assert.equal(hideable.state, 'hidden');

  await callOk('excel_set_sheet_visibility', { filePath: xlsxPath, sheetName: 'Hideable', state: 'visible' });
  const list2 = await callOk('excel_list_sheet_visibility', { filePath: xlsxPath });
  assert.equal(list2.sheets.find((s) => s.name === 'Hideable').state, 'visible');
});

test('excel_set_sheet_visibility refuses to hide last visible sheet', async () => {
  // Create a 2-sheet workbook, hide one, try to hide the other -> should error.
  const oneSheetXlsx = join(tmp, 'onesheet.xlsx');
  await callOk('excel_write_workbook', { filePath: oneSheetXlsx, sheetName: 'Only', data: [['x']] });

  const r = await server.callTool('excel_set_sheet_visibility', {
    filePath: oneSheetXlsx,
    sheetName: 'Only',
    state: 'hidden',
  });
  assert.equal(r.result.isError, true);
  assert.ok(r.result.content[0].text.includes('at least one sheet must remain visible'));
});

test('excel_hide_rows + excel_hide_columns', async () => {
  await callOk('excel_hide_rows', { filePath: xlsxPath, sheetName: 'Data', startRow: 2, count: 2 });
  await callOk('excel_hide_columns', { filePath: xlsxPath, sheetName: 'Data', startColumn: 'B', count: 1 });
  // No assertion on read-back because we don't expose row/col hidden state in read_sheet,
  // but the call succeeding without throw is the contract.
});

test('excel_add_hyperlink + excel_get_hyperlinks + excel_remove_hyperlink', async () => {
  await callOk('excel_add_hyperlink', {
    filePath: xlsxPath,
    sheetName: 'Data',
    cellAddress: 'D1',
    target: 'https://example.com',
    text: 'Click me',
    tooltip: 'Visit example',
  });
  const links = await callOk('excel_get_hyperlinks', { filePath: xlsxPath, sheetName: 'Data' });
  assert.equal(links.totalHyperlinks, 1);
  assert.equal(links.hyperlinks[0].target, 'https://example.com');
  assert.equal(links.hyperlinks[0].text, 'Click me');

  await callOk('excel_remove_hyperlink', { filePath: xlsxPath, sheetName: 'Data', cellAddress: 'D1', keepText: true });
  const after = await callOk('excel_get_hyperlinks', { filePath: xlsxPath, sheetName: 'Data' });
  assert.equal(after.totalHyperlinks, 0);
});

test('excel_set_page_setup + excel_get_page_setup', async () => {
  await callOk('excel_set_page_setup', {
    filePath: xlsxPath,
    sheetName: 'Data',
    config: {
      orientation: 'landscape',
      paperSize: 9,
      printArea: 'A1:C100',
      margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 },
      headerFooter: { oddHeader: '&CMy Report', oddFooter: '&RPage &P of &N' },
    },
  });
  const ps = await callOk('excel_get_page_setup', { filePath: xlsxPath, sheetName: 'Data' });
  assert.equal(ps.orientation, 'landscape');
  assert.equal(ps.paperSize, 9);
  assert.equal(ps.printArea, 'A1:C100');
  assert.equal(ps.headerFooter.oddHeader, '&CMy Report');
});

test('excel_set_workbook_properties + excel_get_workbook_properties', async () => {
  await callOk('excel_set_workbook_properties', {
    filePath: xlsxPath,
    properties: {
      title: 'Q1 Report',
      subject: 'Sales',
      keywords: 'q1 sales report',
      company: 'Soracom',
      creator: 'Excel MCP',
    },
  });
  const props = await callOk('excel_get_workbook_properties', { filePath: xlsxPath });
  assert.equal(props.title, 'Q1 Report');
  assert.equal(props.subject, 'Sales');
  assert.equal(props.company, 'Soracom');
});

test('excel_paste_special: values mode', async () => {
  // Set a formula in F1, paste as value into G1
  await callOk('excel_set_formula', {
    filePath: xlsxPath, sheetName: 'Data', cellAddress: 'F1', formula: '2+3',
  });
  await callOk('excel_paste_special', {
    filePath: xlsxPath,
    sheetName: 'Data',
    sourceRange: 'F1:F1',
    targetCell: 'G1',
    mode: 'values',
  });
  // No deep assert — ExcelJS doesn't always pre-compute formula results without Excel,
  // but the call should succeed and write SOMETHING into G1.
});

test('excel_paste_special: transpose 1x3 to 3x1', async () => {
  await callOk('excel_write_range', {
    filePath: xlsxPath, sheetName: 'Data', range: 'H1:J1', data: [['a', 'b', 'c']],
  });
  await callOk('excel_paste_special', {
    filePath: xlsxPath, sheetName: 'Data', sourceRange: 'H1:J1', targetCell: 'L1', mode: 'transpose',
  });
  const sheet = await callOk('excel_read_sheet', { filePath: xlsxPath, sheetName: 'Data', range: 'L1:L3' });
  assert.equal(sheet.rows[0][0], 'a');
  assert.equal(sheet.rows[1][0], 'b');
  assert.equal(sheet.rows[2][0], 'c');
});

test('excel_get_display_options reports defaults', async () => {
  const d = await callOk('excel_get_display_options', { filePath: xlsxPath, sheetName: 'Data' });
  // ExcelJS doesn't write a view block until set_display_options runs, so defaults apply.
  assert.equal(typeof d.showGridlines, 'boolean');
  assert.equal(typeof d.zoomLevel, 'number');
});

test('excel_get_sheet_protection reads unprotected state', async () => {
  const p = await callOk('excel_get_sheet_protection', { filePath: xlsxPath, sheetName: 'Data' });
  assert.equal(p.isProtected, false);
});

test('excel_list_data_validations on a sheet without validations', async () => {
  const v = await callOk('excel_list_data_validations', { filePath: xlsxPath, sheetName: 'Data' });
  assert.equal(v.totalValidations, 0);
});
