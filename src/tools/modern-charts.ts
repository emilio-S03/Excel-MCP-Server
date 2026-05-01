/**
 * Modern Excel chart types (waterfall, funnel, treemap, sunburst,
 * histogram, box-whisker) and combo charts (mixed column + line on
 * the same plot, with optional secondary value axis).
 *
 * Live mode only — these chart types require the Excel desktop COM
 * interface (Excel 2016+ on Windows). On macOS / Linux we throw a
 * platform-aware error from `platform-errors.ts` that points users
 * at Office Scripts (the supported alternative for modern chart types
 * on those platforms).
 *
 * On Windows we delegate to the PowerShell COM helpers in
 * `excel-powershell.ts`, which know how to:
 *   - assemble a script using `buildPreamble` / `wrapWithCleanup`
 *   - retry through `execPowerShellWithRetry`
 *   - apply the 4-tier series binding fallback for tricky ranges
 *
 * After the COM call succeeds we trigger a `saveFileLive` so the
 * change is persisted, mirroring the behaviour of `createChart` in
 * `charts.ts`.
 */
import { platform } from 'os';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  saveFileLive,
} from './excel-live.js';
import {
  createModernChartViaPowerShell,
  createComboChartViaPowerShell,
} from './excel-powershell.js';
import { winOnlyFileModeAlt } from './platform-errors.js';
import { parseRange } from './helpers.js';
import { ERROR_MESSAGES } from '../constants.js';

const IS_WINDOWS = platform() === 'win32';

export type ModernChartType =
  | 'waterfall'
  | 'funnel'
  | 'treemap'
  | 'sunburst'
  | 'histogram'
  | 'boxWhisker';

export interface CreateModernChartOptions {
  chartType: ModernChartType;
  dataRange: string;
  position: string;
  title?: string;
  dataSheetName?: string;
  createBackup?: boolean;
}

export interface ComboSeriesSpec {
  dataRange: string;
  type: 'column' | 'line';
  color?: string;
  useSecondaryAxis?: boolean;
}

export interface CreateComboChartOptions {
  primarySeries: ComboSeriesSpec;
  secondarySeries: ComboSeriesSpec;
  position: string;
  title?: string;
  createBackup?: boolean;
}

/**
 * Create a modern chart type (waterfall, funnel, treemap, sunburst,
 * histogram, box-whisker). Windows COM only — live mode required.
 */
export async function createModernChart(
  filePath: string,
  sheetName: string,
  options: CreateModernChartOptions
): Promise<string> {
  // Validate range early so the error surface is consistent with
  // excel_create_chart.
  parseRange(options.dataRange);

  if (!IS_WINDOWS) {
    throw winOnlyFileModeAlt(
      'excel_create_modern_chart',
      `${options.chartType} chart creation`,
      'Modern chart types (waterfall/funnel/treemap/sunburst/histogram/boxWhisker) are only available via the Excel desktop COM API on Windows. ' +
      'On macOS, use Microsoft Office Scripts in Excel for the Web (https://learn.microsoft.com/office/dev/scripts/) — they expose these chart types via the JS Excel object model. ' +
      'No file-mode equivalent yet'
    );
  }

  // Live-mode preflight — Excel must be running with the file open.
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  const diagnostics = await createModernChartViaPowerShell(
    filePath,
    sheetName,
    options.chartType,
    options.dataRange,
    options.position,
    options.title,
    options.dataSheetName
  );
  await saveFileLive(filePath);

  // Diagnostics format: "tier=...|seriesCount=...|rows=...|cols=...|chartType=..."
  const diag: Record<string, string> = {};
  for (const part of diagnostics.split('|')) {
    const [k, v] = part.split('=');
    if (k && v) diag[k.trim()] = v.trim();
  }

  return JSON.stringify(
    {
      success: true,
      message: `Modern ${options.chartType} chart created at ${options.position}`,
      chartType: options.chartType,
      dataRange: options.dataRange,
      dataSheetName: options.dataSheetName || sheetName,
      position: options.position,
      title: options.title,
      method: 'live',
      bindingTier: diag['tier'] || 'unknown',
      seriesCount: parseInt(diag['seriesCount'] || '0'),
      dataRows: parseInt(diag['rows'] || '0'),
      dataCols: parseInt(diag['cols'] || '0'),
      note: `Modern chart bound via ${diag['tier'] || 'unknown'} tier. Use excel_style_chart to customize colors, axes, legend.`,
    },
    null,
    2
  );
}

/**
 * Create a combo chart (column + line, optional secondary axis).
 * Windows COM only — live mode required.
 */
export async function createComboChart(
  filePath: string,
  sheetName: string,
  options: CreateComboChartOptions
): Promise<string> {
  parseRange(options.primarySeries.dataRange);
  parseRange(options.secondarySeries.dataRange);

  if (!IS_WINDOWS) {
    throw winOnlyFileModeAlt(
      'excel_create_combo_chart',
      'combo (column + line) chart creation',
      'Combo charts with mixed series types and secondary axes require the Excel desktop COM API on Windows. ' +
      'On macOS, use Microsoft Office Scripts (https://learn.microsoft.com/office/dev/scripts/) — they expose combo chart construction via the JS Excel object model. ' +
      'No file-mode equivalent yet'
    );
  }

  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  const diagnostics = await createComboChartViaPowerShell(filePath, sheetName, {
    primarySeries: options.primarySeries,
    secondarySeries: options.secondarySeries,
    position: options.position,
    title: options.title,
  });
  await saveFileLive(filePath);

  const diag: Record<string, string> = {};
  for (const part of diagnostics.split('|')) {
    const [k, v] = part.split('=');
    if (k && v) diag[k.trim()] = v.trim();
  }

  return JSON.stringify(
    {
      success: true,
      message: `Combo chart (primary=${options.primarySeries.type}, secondary=${options.secondarySeries.type}) created at ${options.position}`,
      primarySeries: options.primarySeries,
      secondarySeries: options.secondarySeries,
      position: options.position,
      title: options.title,
      method: 'live',
      bindingTier: diag['tier'] || 'unknown',
      seriesCount: parseInt(diag['seriesCount'] || '0'),
      primaryCount: parseInt(diag['primaryCount'] || '0'),
      secondaryAxis: diag['secondaryAxis'] === 'true',
      note: 'Combo chart created. Use excel_style_chart to fine-tune per-series formatting.',
    },
    null,
    2
  );
}
