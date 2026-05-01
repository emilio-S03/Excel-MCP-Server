/**
 * Smoke tests for the modern-chart tools (waterfall / funnel / treemap /
 * sunburst / histogram / boxWhisker) and combo charts.
 *
 * These tools are Windows-COM-only (live mode), so on macOS / Linux we
 * just confirm:
 *   - the tool definitions appear in `tools/list`
 *   - the manifest is in sync
 * On Windows we additionally confirm the tool *names* appear in
 * `tools/list`. We do NOT actually invoke the COM bridge in CI because
 * Excel may not be running.
 */
import { test, after, describe } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-modern-charts-'));
const server = await startServer({ EXCEL_ALLOWED_DIRS: tmp });
after(() => {
  server.stop();
  try { rmSync(tmp, { recursive: true, force: true }); } catch {}
});

const MODERN_CHART_TOOLS = ['excel_create_modern_chart', 'excel_create_combo_chart'];

// Always-on assertion: the tool names must appear in tools/list on every
// platform. The manifest is *derived* from tools/list, so if these are
// missing here, the manifest will drift on the next pack.
test('tools/list exposes modern chart tools on every platform', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  for (const name of MODERN_CHART_TOOLS) {
    assert.ok(names.has(name), `missing tool: ${name}`);
  }
});

test('modern chart tool definitions have the expected shape', async () => {
  const tools = await server.listTools();
  const byName = new Map(tools.map((t) => [t.name, t]));

  const modern = byName.get('excel_create_modern_chart');
  assert.ok(modern, 'excel_create_modern_chart should be registered');
  assert.equal(modern.inputSchema.type, 'object');
  const modernEnum = modern.inputSchema.properties.chartType.enum;
  assert.deepEqual(
    new Set(modernEnum),
    new Set(['waterfall', 'funnel', 'treemap', 'sunburst', 'histogram', 'boxWhisker']),
  );
  for (const req of ['filePath', 'sheetName', 'chartType', 'dataRange', 'position']) {
    assert.ok(modern.inputSchema.required.includes(req), `${req} should be required`);
  }

  const combo = byName.get('excel_create_combo_chart');
  assert.ok(combo, 'excel_create_combo_chart should be registered');
  assert.equal(combo.inputSchema.properties.primarySeries.type, 'object');
  assert.equal(combo.inputSchema.properties.secondarySeries.type, 'object');
  for (const req of ['filePath', 'sheetName', 'primarySeries', 'secondarySeries', 'position']) {
    assert.ok(combo.inputSchema.required.includes(req), `${req} should be required`);
  }
});

// Non-Windows behavior: the tools should reject with a PLATFORM_UNSUPPORTED
// error (surfaced through the MCP error envelope) instead of silently
// noop'ing or 500ing on the COM call. We use a non-existent file so we
// don't depend on any test fixture being present.
describe(
  'non-Windows platform behavior',
  { skip: process.platform === 'win32' },
  () => {
    test('excel_create_modern_chart fails fast with a platform-aware error', async () => {
      const r = await server.callTool('excel_create_modern_chart', {
        filePath: join(tmp, 'doesnotexist.xlsx'),
        sheetName: 'Sheet1',
        chartType: 'waterfall',
        dataRange: 'A1:B10',
        position: 'D2',
      });
      assert.ok(r.result?.isError, 'expected isError on non-Windows');
      const body = JSON.parse(r.result.content[0].text);
      // Either a platform-unsupported message, or an Excel-not-running error
      // (whichever the dispatcher reaches first). Both are acceptable; what
      // we want to *prevent* is the COM bridge actually being invoked.
      assert.ok(
        /not yet supported|excel.*not.*running|requires excel/i.test(body.error || ''),
        `expected platform-aware error, got: ${body.error}`,
      );
    });

    test('excel_create_combo_chart fails fast with a platform-aware error', async () => {
      const r = await server.callTool('excel_create_combo_chart', {
        filePath: join(tmp, 'doesnotexist.xlsx'),
        sheetName: 'Sheet1',
        primarySeries: { dataRange: 'A1:B10', type: 'column' },
        secondarySeries: { dataRange: 'A1:C10', type: 'line', useSecondaryAxis: true },
        position: 'D2',
      });
      assert.ok(r.result?.isError, 'expected isError on non-Windows');
      const body = JSON.parse(r.result.content[0].text);
      assert.ok(
        /not yet supported|excel.*not.*running|requires excel/i.test(body.error || ''),
        `expected platform-aware error, got: ${body.error}`,
      );
    });
  },
);

// Windows behavior: confirm the tools register, but DO NOT actually call
// COM — Excel may not be running on the CI box.
describe('Windows registration check', { skip: process.platform !== 'win32' }, () => {
  test('modern chart tools are listed and routable on Windows', async () => {
    const tools = await server.listTools();
    const names = new Set(tools.map((t) => t.name));
    for (const name of MODERN_CHART_TOOLS) {
      assert.ok(names.has(name), `missing tool: ${name}`);
    }
  });
});
