/**
 * Platform dispatcher for live Excel editing.
 * Routes to AppleScript on macOS, PowerShell COM on Windows.
 * On unsupported platforms (Linux), detection returns false so
 * the tool files fall back to ExcelJS file-based editing.
 */
import { platform } from 'os';
import * as applescript from './excel-applescript.js';
import * as macExt from './applescript-extended.js';
import * as powershell from './excel-powershell.js';
import {
  winOnlyUseOfficeScript,
  winOnlyFileModeAlt,
  winOnlyNoAlt,
} from './platform-errors.js';

const IS_WINDOWS = platform() === 'win32';
const IS_MAC = platform() === 'darwin';

// ============================================================
// Detection
// ============================================================

export async function isExcelRunningLive(): Promise<boolean> {
  if (IS_MAC) return applescript.isExcelRunning();
  if (IS_WINDOWS) return powershell.isExcelRunningWindows();
  return false;
}

export async function isFileOpenInExcelLive(filePath: string): Promise<boolean> {
  if (IS_MAC) return applescript.isFileOpenInExcel(filePath);
  if (IS_WINDOWS) return powershell.isFileOpenInExcelWindows(filePath);
  return false;
}

// ============================================================
// Cell Operations
// ============================================================

export async function updateCellLive(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  value: string | number
): Promise<void> {
  if (IS_MAC) return applescript.updateCellViaAppleScript(filePath, sheetName, cellAddress, value);
  if (IS_WINDOWS) return powershell.updateCellViaPowerShell(filePath, sheetName, cellAddress, value);
  throw new Error('Live editing not supported on this platform');
}

export async function writeRangeLive(
  filePath: string,
  sheetName: string,
  startCell: string,
  data: (string | number)[][]
): Promise<void> {
  if (IS_MAC) return applescript.writeRangeViaAppleScript(filePath, sheetName, startCell, data);
  if (IS_WINDOWS) return powershell.writeRangeViaPowerShell(filePath, sheetName, startCell, data);
  throw new Error('Live editing not supported on this platform');
}

export async function addRowLive(
  filePath: string,
  sheetName: string,
  rowData: (string | number)[]
): Promise<void> {
  if (IS_MAC) return applescript.addRowViaAppleScript(filePath, sheetName, rowData);
  if (IS_WINDOWS) return powershell.addRowViaPowerShell(filePath, sheetName, rowData);
  throw new Error('Live editing not supported on this platform');
}

export async function setFormulaLive(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  formula: string
): Promise<void> {
  if (IS_MAC) return applescript.setFormulaViaAppleScript(filePath, sheetName, cellAddress, formula);
  if (IS_WINDOWS) return powershell.setFormulaViaPowerShell(filePath, sheetName, cellAddress, formula);
  throw new Error('Live editing not supported on this platform');
}

// ============================================================
// Formatting
// ============================================================

export async function formatCellLive(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  format: {
    fontName?: string;
    fontSize?: number;
    fontBold?: boolean;
    fontItalic?: boolean;
    fontColor?: string;
    fillColor?: string;
    horizontalAlignment?: string;
    verticalAlignment?: string;
  }
): Promise<void> {
  if (IS_MAC) return applescript.formatCellViaAppleScript(filePath, sheetName, cellAddress, format);
  if (IS_WINDOWS) return powershell.formatCellViaPowerShell(filePath, sheetName, cellAddress, format);
  throw new Error('Live editing not supported on this platform');
}

export async function setColumnWidthLive(
  filePath: string,
  sheetName: string,
  column: string | number,
  width: number
): Promise<void> {
  if (IS_MAC) return applescript.setColumnWidthViaAppleScript(filePath, sheetName, column, width);
  if (IS_WINDOWS) return powershell.setColumnWidthViaPowerShell(filePath, sheetName, column, width);
  throw new Error('Live editing not supported on this platform');
}

export async function setRowHeightLive(
  filePath: string,
  sheetName: string,
  row: number,
  height: number
): Promise<void> {
  if (IS_MAC) return applescript.setRowHeightViaAppleScript(filePath, sheetName, row, height);
  if (IS_WINDOWS) return powershell.setRowHeightViaPowerShell(filePath, sheetName, row, height);
  throw new Error('Live editing not supported on this platform');
}

export async function mergeCellsLive(
  filePath: string,
  sheetName: string,
  range: string
): Promise<void> {
  if (IS_MAC) return applescript.mergeCellsViaAppleScript(filePath, sheetName, range);
  if (IS_WINDOWS) return powershell.mergeCellsViaPowerShell(filePath, sheetName, range);
  throw new Error('Live editing not supported on this platform');
}

export async function unmergeCellsLive(
  filePath: string,
  sheetName: string,
  range: string
): Promise<void> {
  if (IS_MAC) return applescript.unmergeCellsViaAppleScript(filePath, sheetName, range);
  if (IS_WINDOWS) return powershell.unmergeCellsViaPowerShell(filePath, sheetName, range);
  throw new Error('Live editing not supported on this platform');
}

// ============================================================
// Sheet Operations
// ============================================================

export async function createSheetLive(
  filePath: string,
  sheetName: string
): Promise<void> {
  if (IS_MAC) return applescript.createSheetViaAppleScript(filePath, sheetName);
  if (IS_WINDOWS) return powershell.createSheetViaPowerShell(filePath, sheetName);
  throw new Error('Live editing not supported on this platform');
}

export async function deleteSheetLive(
  filePath: string,
  sheetName: string
): Promise<void> {
  if (IS_MAC) return applescript.deleteSheetViaAppleScript(filePath, sheetName);
  if (IS_WINDOWS) return powershell.deleteSheetViaPowerShell(filePath, sheetName);
  throw new Error('Live editing not supported on this platform');
}

export async function renameSheetLive(
  filePath: string,
  oldName: string,
  newName: string
): Promise<void> {
  if (IS_MAC) return applescript.renameSheetViaAppleScript(filePath, oldName, newName);
  if (IS_WINDOWS) return powershell.renameSheetViaPowerShell(filePath, oldName, newName);
  throw new Error('Live editing not supported on this platform');
}

// ============================================================
// Row/Column Operations
// ============================================================

export async function deleteRowsLive(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number
): Promise<void> {
  if (IS_MAC) return applescript.deleteRowsViaAppleScript(filePath, sheetName, startRow, count);
  if (IS_WINDOWS) return powershell.deleteRowsViaPowerShell(filePath, sheetName, startRow, count);
  throw new Error('Live editing not supported on this platform');
}

export async function deleteColumnsLive(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number
): Promise<void> {
  if (IS_MAC) return applescript.deleteColumnsViaAppleScript(filePath, sheetName, startColumn, count);
  if (IS_WINDOWS) return powershell.deleteColumnsViaPowerShell(filePath, sheetName, startColumn, count);
  throw new Error('Live editing not supported on this platform');
}

export async function insertRowsLive(
  filePath: string,
  sheetName: string,
  startRow: number,
  count: number
): Promise<void> {
  if (IS_MAC) return applescript.insertRowsViaAppleScript(filePath, sheetName, startRow, count);
  if (IS_WINDOWS) return powershell.insertRowsViaPowerShell(filePath, sheetName, startRow, count);
  throw new Error('Live editing not supported on this platform');
}

export async function insertColumnsLive(
  filePath: string,
  sheetName: string,
  startColumn: string | number,
  count: number
): Promise<void> {
  if (IS_MAC) return applescript.insertColumnsViaAppleScript(filePath, sheetName, startColumn, count);
  if (IS_WINDOWS) return powershell.insertColumnsViaPowerShell(filePath, sheetName, startColumn, count);
  throw new Error('Live editing not supported on this platform');
}

// ============================================================
// Save
// ============================================================

export async function saveFileLive(filePath: string): Promise<void> {
  if (IS_MAC) return applescript.saveFileViaAppleScript(filePath);
  if (IS_WINDOWS) return powershell.saveFileViaPowerShell(filePath);
  throw new Error('Live editing not supported on this platform');
}

// ============================================================
// Comments
// ============================================================

export async function getCommentsLive(
  filePath: string,
  sheetName: string,
  range?: string
): Promise<string> {
  if (IS_WINDOWS) return powershell.getCommentsViaPowerShell(filePath, sheetName, range);
  throw winOnlyFileModeAlt('excel_get_comments', 'comment reading', 'excel_read_sheet (comments are read with cell metadata when reading the sheet directly from disk)');
}

export async function addCommentLive(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  text: string,
  author?: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.addCommentViaPowerShell(filePath, sheetName, cellAddress, text, author);
  throw winOnlyFileModeAlt('excel_add_comment', 'comment writing', 'excel_update_cell (comments cannot be added via file mode yet — Mac AppleScript port pending)');
}

// ============================================================
// Named Ranges
// ============================================================

export async function listNamedRangesLive(filePath: string): Promise<string> {
  if (IS_WINDOWS) return powershell.listNamedRangesViaPowerShell(filePath);
  throw winOnlyFileModeAlt('excel_list_named_ranges', 'named range listing', 'excel_read_workbook (then inspect names from the file directly — Mac AppleScript port pending)');
}

export async function createNamedRangeLive(
  filePath: string,
  name: string,
  sheetName: string,
  range: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.createNamedRangeViaPowerShell(filePath, name, sheetName, range);
  throw winOnlyNoAlt('excel_create_named_range', 'named range creation while Excel is open');
}

export async function deleteNamedRangeLive(
  filePath: string,
  name: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.deleteNamedRangeViaPowerShell(filePath, name);
  throw winOnlyNoAlt('excel_delete_named_range', 'named range deletion while Excel is open');
}

// ============================================================
// Sheet Protection
// ============================================================

export async function setSheetProtectionLive(
  filePath: string,
  sheetName: string,
  protect: boolean,
  password?: string,
  options?: {
    allowInsertRows?: boolean;
    allowInsertColumns?: boolean;
    allowDeleteRows?: boolean;
    allowDeleteColumns?: boolean;
    allowSort?: boolean;
    allowAutoFilter?: boolean;
    allowFormatCells?: boolean;
    allowFormatColumns?: boolean;
    allowFormatRows?: boolean;
  }
): Promise<void> {
  if (IS_WINDOWS) return powershell.setSheetProtectionViaPowerShell(filePath, sheetName, protect, password, options);
  if (IS_MAC) return macExt.setSheetProtectionViaAppleScript(filePath, sheetName, protect, password);
  throw winOnlyNoAlt('excel_set_sheet_protection', 'live sheet protection');
}

// ============================================================
// Data Validation
// ============================================================

export async function setDataValidationLive(
  filePath: string,
  sheetName: string,
  range: string,
  validationType: string,
  formula1: string,
  operator?: string,
  formula2?: string,
  showErrorMessage?: boolean,
  errorTitle?: string,
  errorMessage?: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.setDataValidationViaPowerShell(filePath, sheetName, range, validationType, formula1, operator, formula2, showErrorMessage, errorTitle, errorMessage);
  throw winOnlyNoAlt('excel_set_data_validation', 'live data validation rules');
}

// ============================================================
// Calculation Control (COM-only)
// ============================================================

export async function triggerRecalculationLive(
  filePath: string,
  fullRecalc: boolean = false
): Promise<void> {
  if (IS_WINDOWS) return powershell.triggerRecalculationViaPowerShell(filePath, fullRecalc);
  if (IS_MAC) return macExt.triggerRecalculationViaAppleScript(filePath, fullRecalc);
  throw winOnlyNoAlt('excel_trigger_recalculation', 'live recalculation trigger');
}

export async function getCalculationModeLive(filePath: string): Promise<string> {
  if (IS_WINDOWS) return powershell.getCalculationModeViaPowerShell(filePath);
  if (IS_MAC) return macExt.getCalculationModeViaAppleScript(filePath);
  throw winOnlyNoAlt('excel_get_calculation_mode', 'reading the live calculation mode');
}

export async function setCalculationModeLive(
  filePath: string,
  mode: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.setCalculationModeViaPowerShell(filePath, mode);
  if (IS_MAC) return macExt.setCalculationModeViaAppleScript(filePath, mode);
  throw winOnlyNoAlt('excel_set_calculation_mode', 'changing the live calculation mode');
}

// ============================================================
// Screenshot (COM-only)
// ============================================================

export async function captureScreenshotLive(
  filePath: string,
  sheetName: string,
  outputPath: string,
  range?: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.captureScreenshotViaPowerShell(filePath, sheetName, outputPath, range);
  if (IS_MAC) return macExt.captureScreenshotViaAppleScript(filePath, sheetName, outputPath, range);
  throw winOnlyNoAlt('excel_capture_screenshot', 'live range capture');
}

// ============================================================
// VBA Macros (COM-only)
// ============================================================

export async function runVbaMacroLive(
  filePath: string,
  macroName: string,
  args: any[] = []
): Promise<string> {
  if (IS_WINDOWS) return powershell.runVbaMacroViaPowerShell(filePath, macroName, args);
  throw winOnlyUseOfficeScript('excel_run_vba_macro', 'VBA macro execution');
}

export async function getVbaCodeLive(
  filePath: string,
  moduleName: string
): Promise<string> {
  if (IS_WINDOWS) return powershell.getVbaCodeViaPowerShell(filePath, moduleName);
  throw winOnlyUseOfficeScript('excel_get_vba_code', 'reading VBA module code');
}

export async function setVbaCodeLive(
  filePath: string,
  moduleName: string,
  code: string,
  createModule: boolean = false,
  appendMode: boolean = false
): Promise<void> {
  if (IS_WINDOWS) return powershell.setVbaCodeViaPowerShell(filePath, moduleName, code, createModule, appendMode);
  throw winOnlyUseOfficeScript('excel_set_vba_code', 'writing VBA module code');
}

// ============================================================
// VBA Trust Access (Registry)
// ============================================================

export async function checkVbaTrustLive(): Promise<string> {
  if (IS_WINDOWS) return powershell.checkVbaTrustViaPowerShell();
  throw winOnlyUseOfficeScript('excel_check_vba_trust', 'reading the VBA Trust Center setting (Windows-only registry path)');
}

export async function enableVbaTrustLive(enable: boolean): Promise<string> {
  if (IS_WINDOWS) return powershell.enableVbaTrustViaPowerShell(enable);
  throw winOnlyUseOfficeScript('excel_enable_vba_trust', 'modifying the VBA Trust Center setting (Windows-only registry path)');
}

// ============================================================
// Diagnosis (Connection & Accessibility)
// ============================================================

export async function diagnoseConnectionLive(filePath?: string): Promise<string> {
  if (IS_WINDOWS) return powershell.diagnoseConnectionViaPowerShell(filePath);
  throw winOnlyFileModeAlt('excel_diagnose_connection', 'COM connection diagnostics', 'excel_check_environment (the cross-platform capability probe — covers Excel install, running state, Automation permission, sandbox dirs)');
}

// ============================================================
// Power Query (COM-only)
// ============================================================

export async function listPowerQueriesLive(filePath: string): Promise<string> {
  if (IS_WINDOWS) return powershell.listPowerQueriesViaPowerShell(filePath);
  throw winOnlyNoAlt('excel_list_power_queries', 'Power Query enumeration (no AppleScript or Mac equivalent)');
}

export async function runPowerQueryLive(
  filePath: string,
  queryName: string,
  formula: string,
  refreshOnly: boolean = false
): Promise<void> {
  if (IS_WINDOWS) return powershell.runPowerQueryViaPowerShell(filePath, queryName, formula, refreshOnly);
  throw winOnlyNoAlt('excel_run_power_query', 'Power Query (M language) execution');
}

// ============================================================
// Chart (Real COM Chart)
// ============================================================

export async function createChartLive(
  filePath: string,
  sheetName: string,
  chartType: string,
  dataRange: string,
  position: string,
  title?: string,
  showLegend: boolean = true,
  dataSheetName?: string
): Promise<string> {
  if (IS_WINDOWS) return powershell.createChartViaPowerShell(filePath, sheetName, chartType, dataRange, position, title, showLegend, dataSheetName);
  throw winOnlyFileModeAlt('excel_create_chart', 'live native chart creation', 'on Mac, charts can be created via the file-mode tool path (current ExcelJS chart support is limited to placeholder shapes — full Mac AppleScript chart implementation is in v3.1)');
}

// ============================================================
// Pivot Table (Real COM Pivot)
// ============================================================

export async function createPivotTableLive(
  filePath: string,
  sourceSheetName: string,
  sourceRange: string,
  targetSheetName: string,
  targetCell: string,
  rows: string[],
  values: Array<{ field: string; aggregation: string }>
): Promise<void> {
  if (IS_WINDOWS) return powershell.createPivotTableViaPowerShell(filePath, sourceSheetName, sourceRange, targetSheetName, targetCell, rows, values);
  throw winOnlyFileModeAlt('excel_create_pivot_table', 'live pivot table creation', 'the file-mode pivot path (computes aggregations and writes static values — no refreshable PivotTable yet on Mac)');
}

// ============================================================
// Table (COM)
// ============================================================

export async function createTableLive(
  filePath: string,
  sheetName: string,
  range: string,
  tableName: string,
  tableStyle: string = 'TableStyleMedium2'
): Promise<void> {
  if (IS_WINDOWS) return powershell.createTableViaPowerShell(filePath, sheetName, range, tableName, tableStyle);
  throw winOnlyNoAlt('excel_create_table', 'live ListObject (Excel Table) creation');
}

// ============================================================
// Conditional Formatting (COM)
// ============================================================

export async function applyConditionalFormatLive(
  filePath: string,
  sheetName: string,
  range: string,
  ruleType: string,
  condition?: {
    operator?: string;
    value?: any;
    value2?: any;
  },
  style?: {
    font?: { color?: string; bold?: boolean };
    fill?: { fgColor?: string };
  },
  colorScale?: {
    minColor?: string;
    midColor?: string;
    maxColor?: string;
  }
): Promise<void> {
  if (IS_WINDOWS) return powershell.applyConditionalFormatViaPowerShell(filePath, sheetName, range, ruleType, condition, style, colorScale);
  throw winOnlyFileModeAlt('excel_apply_conditional_format', 'live conditional formatting', 'the file-mode equivalent (closes the file, applies via ExcelJS, reopen to see changes)');
}

// ============================================================
// Batch Format (COM)
// ============================================================

export async function batchFormatLive(
  filePath: string,
  sheetName: string,
  operations: Array<{
    range: string;
    merge?: boolean;
    unmerge?: boolean;
    value?: string | number;
    fontName?: string;
    fontSize?: number;
    fontBold?: boolean;
    fontItalic?: boolean;
    fontColor?: string;
    fillColor?: string;
    horizontalAlignment?: string;
    verticalAlignment?: string;
    numberFormat?: string;
    columnWidth?: number;
    rowHeight?: number;
    borderStyle?: string;
    borderColor?: string;
    wrapText?: boolean;
    autoFit?: boolean;
  }>
): Promise<void> {
  if (IS_WINDOWS) return powershell.batchFormatViaPowerShell(filePath, sheetName, operations);
  throw winOnlyFileModeAlt('excel_batch_format', 'live batch formatting', 'individual file-mode tools (excel_format_cell, excel_set_column_width) which work on Mac without Excel running');
}

// ============================================================
// Display Options (COM)
// ============================================================

export async function setDisplayOptionsLive(
  filePath: string,
  sheetName?: string,
  showGridlines?: boolean,
  showRowColumnHeaders?: boolean,
  zoomLevel?: number,
  freezePaneCell?: string,
  tabColor?: string
): Promise<void> {
  if (IS_WINDOWS) return powershell.setDisplayOptionsViaPowerShell(filePath, sheetName, showGridlines, showRowColumnHeaders, zoomLevel, freezePaneCell, tabColor);
  if (IS_MAC) return macExt.setDisplayOptionsViaAppleScript(filePath, sheetName, showGridlines, showRowColumnHeaders, zoomLevel, freezePaneCell, tabColor);
  throw winOnlyNoAlt('excel_set_display_options', 'live display options');
}

// ============================================================
// Shapes (COM)
// ============================================================

export async function addShapeLive(
  filePath: string,
  sheetName: string,
  config: {
    shapeType: string;
    left: number;
    top: number;
    width: number;
    height: number;
    name?: string;
    fill?: {
      color?: string;
      gradient?: { color1: string; color2: string; direction?: string };
      transparency?: number;
    };
    line?: { color?: string; weight?: number; visible?: boolean };
    shadow?: { visible?: boolean; color?: string; offsetX?: number; offsetY?: number; blur?: number; transparency?: number };
    text?: { value: string; fontName?: string; fontSize?: number; fontBold?: boolean; fontColor?: string; horizontalAlignment?: string; verticalAlignment?: string; autoSize?: string };
  }
): Promise<string> {
  if (IS_WINDOWS) return powershell.addShapeViaPowerShell(filePath, sheetName, config);
  throw winOnlyFileModeAlt('excel_add_shape', 'live shape insertion (used for dashboard cards)', 'excel_add_image (file-mode — embed an image of the dashboard card design instead) until the AppleScript shape port lands in v3.1');
}

// ============================================================
// Chart Styling (COM)
// ============================================================

export async function styleChartLive(
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
): Promise<void> {
  if (IS_WINDOWS) return powershell.styleChartViaPowerShell(filePath, sheetName, chartIndex, chartName, config);
  throw winOnlyFileModeAlt('excel_style_chart', 'live native chart restyling', 'excel_create_chart in file-mode (limited styling) — full Mac chart styling lands in v3.1');
}

// ============================================================
// Live-mode INSPECTION (read existing charts/pivots/shapes)
// ============================================================

export async function listChartsLive(filePath: string, sheetName?: string): Promise<string> {
  if (IS_WINDOWS) return powershell.listChartsViaPowerShell(filePath, sheetName);
  if (IS_MAC) return macExt.listChartsViaAppleScript(filePath, sheetName);
  throw winOnlyFileModeAlt('excel_list_charts', 'live chart enumeration', 'no file-mode equivalent — ExcelJS does not preserve real chart definitions, so inspection requires Excel to be running');
}

export async function getChartLive(
  filePath: string,
  sheetName: string,
  chartIndex?: number,
  chartName?: string
): Promise<string> {
  if (IS_WINDOWS) return powershell.getChartViaPowerShell(filePath, sheetName, chartIndex, chartName);
  if (IS_MAC) return macExt.getChartViaAppleScript(filePath, sheetName, chartIndex, chartName);
  throw winOnlyFileModeAlt('excel_get_chart', 'live chart detail inspection', 'no file-mode equivalent — ExcelJS does not preserve real chart definitions, so inspection requires Excel to be running');
}

export async function listPivotTablesLive(filePath: string, sheetName?: string): Promise<string> {
  if (IS_WINDOWS) return powershell.listPivotTablesViaPowerShell(filePath, sheetName);
  if (IS_MAC) return macExt.listPivotTablesViaAppleScript(filePath, sheetName);
  throw winOnlyFileModeAlt('excel_list_pivot_tables', 'live pivot table enumeration', 'no file-mode equivalent — ExcelJS does not preserve real PivotTable definitions');
}

export async function listShapesLive(filePath: string, sheetName?: string): Promise<string> {
  if (IS_WINDOWS) return powershell.listShapesViaPowerShell(filePath, sheetName);
  if (IS_MAC) return macExt.listShapesViaAppleScript(filePath, sheetName);
  throw winOnlyFileModeAlt('excel_list_shapes', 'live shape enumeration', 'no file-mode equivalent — ExcelJS does not preserve shape attributes faithfully');
}
