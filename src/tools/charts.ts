import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';
import { isExcelRunningLive, isFileOpenInExcelLive, createChartLive, styleChartLive, saveFileLive } from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

export async function createChart(
  filePath: string,
  sheetName: string,
  chartType: 'line' | 'bar' | 'column' | 'pie' | 'scatter' | 'area',
  dataRange: string,
  position: string,
  title?: string,
  showLegend: boolean = true,
  createBackup: boolean = false,
  dataSheetName?: string
): Promise<string> {
  // Validate the data range
  parseRange(dataRange);

  // Check if Excel is running and file is open — use real COM chart if so
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    const diagnostics = await createChartLive(filePath, sheetName, chartType, dataRange, position, title, showLegend, dataSheetName);
    await saveFileLive(filePath);

    // Parse diagnostics: "tier=SetSourceData|seriesCount=1|rows=15|cols=2"
    const diag: Record<string, string> = {};
    for (const part of diagnostics.split('|')) {
      const [k, v] = part.split('=');
      if (k && v) diag[k.trim()] = v.trim();
    }

    return JSON.stringify({
      success: true,
      message: `Real ${chartType} chart created at ${position}`,
      chartType,
      dataRange,
      dataSheetName: dataSheetName || sheetName,
      position,
      title,
      method: 'live',
      bindingTier: diag['tier'] || 'unknown',
      seriesCount: parseInt(diag['seriesCount'] || '0'),
      dataRows: parseInt(diag['rows'] || '0'),
      dataCols: parseInt(diag['cols'] || '0'),
      note: `Chart data bound via ${diag['tier'] || 'unknown'} tier. ${diag['seriesCount'] || 0} series from ${diag['rows'] || '?'} rows x ${diag['cols'] || '?'} cols.`,
    }, null, 2);
  }

  // ExcelJS fallback — placeholder only
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const chartInfo = {
    type: chartType,
    dataRange,
    position,
    title,
    showLegend,
    note: 'Chart placeholder created. Open file in Excel for full chart support.'
  };

  const posCell = sheet.getCell(position);
  posCell.value = title || `${chartType.toUpperCase()} Chart`;
  posCell.note = JSON.stringify(chartInfo, null, 2);

  posCell.font = { bold: true, size: 14 };
  posCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE7E6E6' },
  };

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Chart placeholder created at ${position}`,
    note: 'ExcelJS has limited native chart support. Open the file in Excel and use excel_create_chart with Excel running for real charts.',
    chartType,
    dataRange,
    position,
    title,
    method: 'exceljs',
  }, null, 2);
}

export async function styleChart(
  filePath: string,
  sheetName: string,
  chartIndex: number | undefined,
  chartName: string | undefined,
  config: {
    series?: Array<{
      index: number;
      color?: string;
      lineWeight?: number;
      markerStyle?: string;
      markerSize?: number;
      dataLabels?: { show: boolean; numberFormat?: string; fontSize?: number; fontColor?: string; position?: string };
    }>;
    axes?: {
      category?: { visible?: boolean; numberFormat?: string; fontSize?: number; fontColor?: string; labelRotation?: number };
      value?: { visible?: boolean; numberFormat?: string; fontSize?: number; fontColor?: string; min?: number; max?: number; gridlines?: boolean };
    };
    chartArea?: { fillColor?: string; borderVisible?: boolean };
    plotArea?: { fillColor?: string };
    legend?: { visible: boolean; position?: string; fontSize?: number; fontColor?: string };
    title?: { text?: string; visible?: boolean; fontSize?: number; fontColor?: string };
    width?: number;
    height?: number;
  }
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  await styleChartLive(filePath, sheetName, chartIndex, chartName, config);

  return JSON.stringify({
    success: true,
    message: `Chart styled on sheet "${sheetName}"`,
    chartIndex,
    chartName,
    method: 'live',
    note: 'Chart styling applied via COM. Changes visible immediately in Excel.',
  }, null, 2);
}
