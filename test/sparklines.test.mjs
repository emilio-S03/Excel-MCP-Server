import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { startServer } from './harness.mjs';

const tmp = mkdtempSync(join(tmpdir(), 'excel-mcp-sparklines-'));
const xlsxPath = join(tmp, 'sparklines.xlsx');

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

async function readSheetXml(filePath, sheetName) {
  const zip = await JSZip.loadAsync(readFileSync(filePath));
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const sheetEntries = [...wbXml.matchAll(/<sheet\b[^>]*?\/>/g)];
  let rid = null;
  for (const m of sheetEntries) {
    const nameMatch = m[0].match(/\bname="([^"]+)"/);
    const ridMatch = m[0].match(/\b(?:r:id|r:Id)="([^"]+)"/);
    if (nameMatch && ridMatch && nameMatch[1] === sheetName) { rid = ridMatch[1]; break; }
  }
  assert.ok(rid, `sheet ${sheetName} not found in workbook.xml`);
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relEntries = [...relsXml.matchAll(/<Relationship\b[^>]*?\/>/g)];
  let target = null;
  for (const m of relEntries) {
    const idMatch = m[0].match(/\bId="([^"]+)"/);
    const tgtMatch = m[0].match(/\bTarget="([^"]+)"/);
    if (idMatch && tgtMatch && idMatch[1] === rid) { target = tgtMatch[1]; break; }
  }
  assert.ok(target, `relationship ${rid} not resolvable`);
  const path = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
  return zip.file(path).async('string');
}

test('excel_add_sparkline injects extLst with sparkline group', async () => {
  // Bootstrap — write a workbook with numeric data in A1:A10 on Sheet1.
  const data = [[5], [12], [-3], [22], [17], [9], [-8], [14], [25], [4]];
  await callOk('excel_write_workbook', {
    filePath: xlsxPath,
    sheetName: 'Sheet1',
    data,
  });

  const result = await callOk('excel_add_sparkline', {
    filePath: xlsxPath,
    sheetName: 'Sheet1',
    type: 'line',
    dataRange: 'A1:A10',
    locationRange: 'B1',
    color: '#376092',
    markers: true,
    high: true,
    low: true,
  });
  assert.equal(result.success, true);
  assert.equal(result.sparklineCount, 1);

  const sheetXml = await readSheetXml(xlsxPath, 'Sheet1');
  assert.ok(/<extLst>/.test(sheetXml), 'expected <extLst> block in sheet xml');
  assert.ok(
    /uri="\{05C60535-1F16-4fd2-B633-F4F36F0B6A02\}"/.test(sheetXml),
    'expected sparkline ext URI',
  );
  assert.ok(/<x14:sparklineGroups\b/.test(sheetXml), 'expected sparklineGroups element');
  assert.ok(/<x14:sparklineGroup\b/.test(sheetXml), 'expected sparklineGroup element');
  assert.ok(/<x14:sparkline>/.test(sheetXml), 'expected sparkline element');
  assert.ok(/<xm:f>Sheet1!A1:A10<\/xm:f>/.test(sheetXml), 'expected qualified data range');
  assert.ok(/<xm:sqref>B1<\/xm:sqref>/.test(sheetXml), 'expected location cell B1');
  assert.ok(/<x14:colorSeries rgb="FF376092"\/>/.test(sheetXml), 'expected series color');
  assert.ok(/markers="1"/.test(sheetXml), 'expected markers="1"');
  assert.ok(/high="1"/.test(sheetXml), 'expected high="1"');
  assert.ok(/low="1"/.test(sheetXml), 'expected low="1"');
});

test('excel_add_sparkline with vertical location range slices by row', async () => {
  // Re-bootstrap with a 5x3 numeric block; one sparkline per row in column E.
  const data = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
    [10, 11, 12],
    [13, 14, 15],
  ];
  await callOk('excel_write_workbook', {
    filePath: xlsxPath,
    sheetName: 'Sheet1',
    data,
  });

  const result = await callOk('excel_add_sparkline', {
    filePath: xlsxPath,
    sheetName: 'Sheet1',
    type: 'column',
    dataRange: 'A1:C5',
    locationRange: 'E1:E5',
  });
  assert.equal(result.sparklineCount, 5);

  const sheetXml = await readSheetXml(xlsxPath, 'Sheet1');
  assert.ok(/<x14:sparkline>/.test(sheetXml));
  // Row 1 should map to A1:C1, row 5 to A5:C5.
  assert.ok(/<xm:f>Sheet1!A1:C1<\/xm:f>/.test(sheetXml), 'expected first row slice');
  assert.ok(/<xm:f>Sheet1!A5:C5<\/xm:f>/.test(sheetXml), 'expected last row slice');
  assert.ok(/<xm:sqref>E1<\/xm:sqref>/.test(sheetXml));
  assert.ok(/<xm:sqref>E5<\/xm:sqref>/.test(sheetXml));
  assert.ok(/type="column"/.test(sheetXml));
});

test('excel_remove_sparklines clears all groups when no locationRange given', async () => {
  const before = await readSheetXml(xlsxPath, 'Sheet1');
  assert.ok(/<x14:sparklineGroup\b/.test(before), 'precondition: sparklines exist');

  const r = await callOk('excel_remove_sparklines', {
    filePath: xlsxPath,
    sheetName: 'Sheet1',
  });
  assert.equal(r.success, true);
  assert.ok(r.removedSparklines >= 1, 'should report at least one removed sparkline');

  const after = await readSheetXml(xlsxPath, 'Sheet1');
  assert.ok(!/<x14:sparklineGroup\b/.test(after), 'no sparkline groups after removal');
});

test('tools/list reports the new sparkline tools', async () => {
  const tools = await server.listTools();
  const names = new Set(tools.map((t) => t.name));
  assert.ok(names.has('excel_add_sparkline'), 'excel_add_sparkline missing');
  assert.ok(names.has('excel_remove_sparklines'), 'excel_remove_sparklines missing');
});

test('excel_check_environment lists the new sparkline tools in fileMode', async () => {
  const r = await callOk('excel_check_environment', {});
  const tools = r.capabilityMatrix?.fileMode?.tools ?? [];
  assert.ok(tools.includes('excel_add_sparkline'), 'excel_add_sparkline missing from fileMode tools');
  assert.ok(tools.includes('excel_remove_sparklines'), 'excel_remove_sparklines missing from fileMode tools');
});
