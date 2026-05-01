import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from './harness.mjs';

test('default sandbox includes Documents/Desktop/Downloads', async () => {
  const server = await startServer({ EXCEL_ALLOWED_DIRS: '' });
  try {
    // Trigger anything so the boot stderr line is flushed
    await server.listTools();
    await new Promise((r) => setTimeout(r, 100));
    const log = server.getStderr();
    assert.ok(log.includes('Documents'));
    assert.ok(log.includes('Desktop'));
    assert.ok(log.includes('Downloads'));
  } finally {
    server.stop();
  }
});

test('EXCEL_ALLOWED_DIRS overrides defaults; outside writes are rejected', async () => {
  const sandbox = mkdtempSync(join(tmpdir(), 'excel-sandbox-'));
  const tmpOutside = mkdtempSync(join(tmpdir(), 'excel-outside-'));
  const server = await startServer({ EXCEL_ALLOWED_DIRS: sandbox });
  try {
    const blocked = await server.callTool('excel_write_workbook', {
      filePath: join(tmpOutside, 'blocked.xlsx'),
      sheetName: 'X',
      data: [[1]],
    });
    assert.equal(blocked.result.isError, true);
    assert.ok(blocked.result.content[0].text.includes('PATH_OUTSIDE_ALLOWED'));

    const ok = await server.callTool('excel_write_workbook', {
      filePath: join(sandbox, 'ok.xlsx'),
      sheetName: 'X',
      data: [[1, 2, 3]],
    });
    assert.equal(ok.result.isError, undefined);
  } finally {
    server.stop();
    try { rmSync(sandbox, { recursive: true, force: true }); } catch {}
    try { rmSync(tmpOutside, { recursive: true, force: true }); } catch {}
  }
});

test('default sandbox blocks system paths', async () => {
  const server = await startServer({});
  try {
    const target = process.platform === 'win32'
      ? 'C:/Windows/Temp/should-not-write.xlsx'
      : '/etc/should-not-write.xlsx';
    const blocked = await server.callTool('excel_write_workbook', {
      filePath: target,
      sheetName: 'X',
      data: [[1]],
    });
    assert.equal(blocked.result.isError, true);
    assert.ok(blocked.result.content[0].text.includes('PATH_OUTSIDE_ALLOWED'));
  } finally {
    server.stop();
  }
});
