/**
 * Live-mode inspection tools — read existing charts, pivots, and shapes
 * from a workbook open in Excel. These complement the create/style tools
 * (excel_create_chart, excel_create_pivot_table, excel_add_shape, excel_style_chart)
 * with a read-only counterpart.
 *
 * Why live-only? File-mode (ExcelJS) does NOT preserve real chart definitions,
 * pivot caches, or shape attributes round-trip. The only reliable way to read
 * these is to ask Excel directly via COM (Windows) or AppleScript (Mac, UNVERIFIED).
 *
 * Tool list:
 *   - excel_list_charts(filePath, sheetName?)
 *   - excel_get_chart(filePath, sheetName, chartIndex|chartName)
 *   - excel_list_pivot_tables(filePath, sheetName?)
 *   - excel_list_shapes(filePath, sheetName?)
 */
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  listChartsLive,
  getChartLive,
  listPivotTablesLive,
  listShapesLive,
} from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

async function ensureLive(filePath: string): Promise<void> {
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
}

function safeParse(raw: string): unknown {
  if (!raw || !raw.trim()) return [];
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}

export async function listCharts(filePath: string, sheetName?: string): Promise<string> {
  await ensureLive(filePath);
  const raw = await listChartsLive(filePath, sheetName);
  const charts = safeParse(raw);
  return JSON.stringify({
    charts: Array.isArray(charts) ? charts : (charts ? [charts] : []),
    sheetFilter: sheetName ?? null,
    method: 'live',
  }, null, 2);
}

export async function getChart(
  filePath: string,
  sheetName: string,
  chartIndex?: number,
  chartName?: string
): Promise<string> {
  if (chartIndex === undefined && !chartName) {
    throw new Error('Either chartIndex or chartName must be provided');
  }
  await ensureLive(filePath);
  const raw = await getChartLive(filePath, sheetName, chartIndex, chartName);
  const chart = safeParse(raw);
  return JSON.stringify({
    chart,
    method: 'live',
  }, null, 2);
}

export async function listPivotTables(filePath: string, sheetName?: string): Promise<string> {
  await ensureLive(filePath);
  const raw = await listPivotTablesLive(filePath, sheetName);
  const pivots = safeParse(raw);
  return JSON.stringify({
    pivotTables: Array.isArray(pivots) ? pivots : (pivots ? [pivots] : []),
    sheetFilter: sheetName ?? null,
    method: 'live',
  }, null, 2);
}

export async function listShapes(filePath: string, sheetName?: string): Promise<string> {
  await ensureLive(filePath);
  const raw = await listShapesLive(filePath, sheetName);
  const shapes = safeParse(raw);
  return JSON.stringify({
    shapes: Array.isArray(shapes) ? shapes : (shapes ? [shapes] : []),
    sheetFilter: sheetName ?? null,
    method: 'live',
  }, null, 2);
}
