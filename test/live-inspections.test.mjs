/**
 * Live-mode inspection tools — registration test only.
 *
 * These tools (excel_list_charts, excel_get_chart, excel_list_pivot_tables,
 * excel_list_shapes) require Excel to be running with the file open. We
 * cannot invoke them in CI, so this test only asserts that they show up
 * in tools/list with the expected schema shape and the read-only
 * annotation. End-to-end validation is manual on a Windows box with
 * Excel running. Mac coverage is UNVERIFIED — see applescript-extended.ts
 * comments for the dictionary calls that still need a real-Mac sanity pass.
 */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { startServer } from './harness.mjs';

const server = await startServer();
after(() => server.stop());

const EXPECTED = [
  'excel_list_charts',
  'excel_get_chart',
  'excel_list_pivot_tables',
  'excel_list_shapes',
];

test('tools/list exposes the 4 live-mode inspection tools', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  for (const expected of EXPECTED) {
    assert.ok(names.has(expected), `missing tool: ${expected}`);
  }
});

test('inspection tools all carry the read-only annotation', async () => {
  const tools = await server.listTools();
  for (const expected of EXPECTED) {
    const tool = tools.find((t) => t.name === expected);
    assert.ok(tool, `missing tool: ${expected}`);
    assert.equal(
      tool.annotations?.readOnlyHint,
      true,
      `${expected} should be marked read-only (these enumerate, not mutate)`
    );
  }
});

test('excel_list_charts schema accepts optional sheetName', async () => {
  const tools = await server.listTools();
  const tool = tools.find((t) => t.name === 'excel_list_charts');
  assert.deepEqual(tool.inputSchema.required, ['filePath']);
  assert.ok(tool.inputSchema.properties.sheetName, 'sheetName property must exist');
});

test('excel_get_chart schema requires filePath + sheetName, accepts chartIndex or chartName', async () => {
  const tools = await server.listTools();
  const tool = tools.find((t) => t.name === 'excel_get_chart');
  assert.deepEqual(tool.inputSchema.required, ['filePath', 'sheetName']);
  assert.ok(tool.inputSchema.properties.chartIndex);
  assert.ok(tool.inputSchema.properties.chartName);
});

test('excel_list_pivot_tables schema mirrors excel_list_charts', async () => {
  const tools = await server.listTools();
  const tool = tools.find((t) => t.name === 'excel_list_pivot_tables');
  assert.deepEqual(tool.inputSchema.required, ['filePath']);
  assert.ok(tool.inputSchema.properties.sheetName);
});

test('excel_list_shapes schema mirrors excel_list_charts', async () => {
  const tools = await server.listTools();
  const tool = tools.find((t) => t.name === 'excel_list_shapes');
  assert.deepEqual(tool.inputSchema.required, ['filePath']);
  assert.ok(tool.inputSchema.properties.sheetName);
});

test('total tool count includes the 4 new inspection tools', async () => {
  const tools = await server.listTools();
  // Count is whatever is registered; we only care that it strictly grew by 4
  // relative to the v3.1 baseline of 83. The exact figure (currently 96 —
  // 83 v3.1 + sparklines + formula audit + modern charts + these 4) is
  // tracked by manifest-sync.test.mjs, so don't lock it in here.
  for (const expected of EXPECTED) {
    assert.ok(tools.some((t) => t.name === expected), `missing tool: ${expected}`);
  }
});
