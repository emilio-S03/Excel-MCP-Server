import { z } from 'zod';

// Common schemas
export const responseFormatSchema = z.enum(['json', 'markdown']).default('json');
export const filePathSchema = z.string().describe('Path to the Excel file');
export const sheetNameSchema = z.string().describe('Name of the sheet');
export const cellAddressSchema = z.string().regex(/^[A-Z]+\d+$/, 'Invalid cell address (e.g., A1, B2)');
export const rangeSchema = z.string().regex(/^[A-Z]+\d+:[A-Z]+\d+$/, 'Invalid range (e.g., A1:D10)');

// Read operations
export const readWorkbookSchema = z.object({
  filePath: filePathSchema,
  responseFormat: responseFormatSchema,
});

export const readSheetSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.optional().describe('Optional range to read (e.g., A1:D10)'),
  responseFormat: responseFormatSchema,
});

export const readRangeSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  responseFormat: responseFormatSchema,
});

export const getCellSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  responseFormat: responseFormatSchema,
});

export const getFormulaSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  responseFormat: responseFormatSchema,
});

// Write operations
export const writeWorkbookSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema.default('Sheet1'),
  data: z.array(z.array(z.any())).describe('2D array of data to write'),
  createBackup: z.boolean().default(false),
});

export const updateCellSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  value: z.any().describe('Value to write to the cell'),
  createBackup: z.boolean().default(false),
});

export const writeRangeSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  data: z.array(z.array(z.any())).describe('2D array of data to write'),
  createBackup: z.boolean().default(false),
});

export const addRowSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  data: z.array(z.any()).describe('Array of values for the new row'),
  createBackup: z.boolean().default(false),
});

export const setFormulaSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  formula: z.string().describe('Excel formula (without = sign)'),
  createBackup: z.boolean().default(false),
});

// Format operations
export const formatCellSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  format: z.object({
    font: z.object({
      name: z.string().optional(),
      size: z.number().optional(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      color: z.string().optional().describe('Hex color code (e.g., FF0000 for red)'),
    }).optional(),
    fill: z.object({
      type: z.literal('pattern'),
      pattern: z.enum(['solid', 'darkVertical', 'darkHorizontal', 'darkGrid']),
      fgColor: z.string().optional().describe('Foreground hex color'),
      bgColor: z.string().optional().describe('Background hex color'),
    }).optional(),
    alignment: z.object({
      horizontal: z.enum(['left', 'center', 'right', 'fill', 'justify']).optional(),
      vertical: z.enum(['top', 'middle', 'bottom']).optional(),
      wrapText: z.boolean().optional(),
    }).optional(),
    border: z.object({
      top: z.object({ style: z.string(), color: z.string().optional() }).optional(),
      left: z.object({ style: z.string(), color: z.string().optional() }).optional(),
      bottom: z.object({ style: z.string(), color: z.string().optional() }).optional(),
      right: z.object({ style: z.string(), color: z.string().optional() }).optional(),
    }).optional(),
    numFmt: z.string().optional().describe('Number format (e.g., "0.00", "$#,##0.00")'),
  }),
  createBackup: z.boolean().default(false),
});

export const setColumnWidthSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  column: z.union([z.string(), z.number()]).describe('Column letter (A, B, C) or number (1, 2, 3)'),
  width: z.number().describe('Width in Excel units (approximately characters)'),
  createBackup: z.boolean().default(false),
});

export const setRowHeightSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  row: z.number().describe('Row number (1-based)'),
  height: z.number().describe('Height in points'),
  createBackup: z.boolean().default(false),
});

export const mergeCellsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  createBackup: z.boolean().default(false),
});

// Sheet management
export const createSheetSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  createBackup: z.boolean().default(false),
});

export const deleteSheetSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  createBackup: z.boolean().default(false),
});

export const renameSheetSchema = z.object({
  filePath: filePathSchema,
  oldName: z.string().describe('Current sheet name'),
  newName: z.string().describe('New sheet name'),
  createBackup: z.boolean().default(false),
});

export const duplicateSheetSchema = z.object({
  filePath: filePathSchema,
  sourceSheetName: z.string().describe('Name of sheet to duplicate'),
  newSheetName: z.string().describe('Name for the duplicated sheet'),
  createBackup: z.boolean().default(false),
});

// Operations
export const deleteRowsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startRow: z.number().describe('Starting row number (1-based)'),
  count: z.number().describe('Number of rows to delete'),
  createBackup: z.boolean().default(false),
});

export const deleteColumnsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startColumn: z.union([z.string(), z.number()]).describe('Starting column (letter or number)'),
  count: z.number().describe('Number of columns to delete'),
  createBackup: z.boolean().default(false),
});

export const copyRangeSchema = z.object({
  filePath: filePathSchema,
  sourceSheetName: sheetNameSchema,
  sourceRange: rangeSchema,
  targetSheetName: sheetNameSchema,
  targetCell: cellAddressSchema.describe('Top-left cell of destination'),
  createBackup: z.boolean().default(false),
});

// Analysis
export const searchValueSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  searchValue: z.any().describe('Value to search for'),
  range: rangeSchema.optional().describe('Optional range to search within'),
  caseSensitive: z.boolean().default(false),
  responseFormat: responseFormatSchema,
});

export const filterRowsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  column: z.union([z.string(), z.number()]).describe('Column to filter by'),
  condition: z.enum(['equals', 'contains', 'greater_than', 'less_than', 'not_empty']),
  value: z.any().optional().describe('Value to compare against (not needed for not_empty)'),
  responseFormat: responseFormatSchema,
});

// Charts
export const createChartSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  chartType: z.enum(['line', 'bar', 'column', 'pie', 'scatter', 'area']),
  dataRange: rangeSchema.describe('Range of data for the chart'),
  dataSheetName: z.string().optional().describe('Sheet containing the data range (if different from sheetName where chart is placed)'),
  position: cellAddressSchema.describe('Top-left cell where chart will be placed'),
  title: z.string().optional().describe('Chart title'),
  showLegend: z.boolean().default(true),
  createBackup: z.boolean().default(false),
});

// Pivot Tables
export const createPivotTableSchema = z.object({
  filePath: filePathSchema,
  sourceSheetName: sheetNameSchema.describe('Sheet containing source data'),
  sourceRange: rangeSchema.describe('Range of source data'),
  targetSheetName: sheetNameSchema.describe('Sheet for pivot table'),
  targetCell: cellAddressSchema.describe('Top-left cell for pivot table'),
  rows: z.array(z.string()).describe('Fields for row labels'),
  columns: z.array(z.string()).optional().describe('Fields for column labels'),
  values: z.array(z.object({
    field: z.string(),
    aggregation: z.enum(['sum', 'count', 'average', 'min', 'max']),
  })).describe('Fields to aggregate'),
  createBackup: z.boolean().default(false),
});

// Tables
export const createTableSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.describe('Range to convert to table'),
  tableName: z.string().describe('Name for the table'),
  tableStyle: z.string().optional().default('TableStyleMedium2').describe('Excel table style'),
  showFirstColumn: z.boolean().default(false),
  showLastColumn: z.boolean().default(false),
  showRowStripes: z.boolean().default(true),
  showColumnStripes: z.boolean().default(false),
  createBackup: z.boolean().default(false),
});

// Validation operations
export const validateFormulaSyntaxSchema = z.object({
  formula: z.string().describe('Formula to validate (without = sign)'),
});

export const validateExcelRangeSchema = z.object({
  range: z.string().describe('Range to validate (e.g., A1:D10)'),
});

export const getDataValidationInfoSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  responseFormat: responseFormatSchema,
});

// Advanced operations
export const insertRowsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startRow: z.number().describe('Row number where to insert (1-based)'),
  count: z.number().describe('Number of rows to insert'),
  createBackup: z.boolean().default(false),
});

export const insertColumnsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  startColumn: z.union([z.string(), z.number()]).describe('Column where to insert (letter or number)'),
  count: z.number().describe('Number of columns to insert'),
  createBackup: z.boolean().default(false),
});

export const unmergeCellsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.describe('Range to unmerge'),
  createBackup: z.boolean().default(false),
});

export const getMergedCellsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  responseFormat: responseFormatSchema,
});

// Conditional formatting
export const applyConditionalFormatSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema,
  ruleType: z.enum(['cellValue', 'colorScale', 'dataBar', 'topBottom']),
  condition: z.object({
    operator: z.enum(['greaterThan', 'lessThan', 'between', 'equal', 'notEqual', 'containsText']).optional(),
    value: z.any().optional(),
    value2: z.any().optional().describe('Second value for "between" operator'),
  }).optional(),
  style: z.object({
    font: z.object({
      color: z.string().optional(),
      bold: z.boolean().optional(),
    }).optional(),
    fill: z.object({
      type: z.literal('pattern'),
      pattern: z.enum(['solid', 'darkVertical', 'darkHorizontal', 'darkGrid']),
      fgColor: z.string().optional(),
    }).optional(),
  }).optional(),
  colorScale: z.object({
    minColor: z.string().optional().default('FFFF0000'),
    midColor: z.string().optional(),
    maxColor: z.string().optional().default('FF00FF00'),
  }).optional().describe('For colorScale type'),
  createBackup: z.boolean().default(false),
});

// ============================================================
// Comments
// ============================================================

export const getCommentsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.optional().describe('Optional range to get comments from (e.g., A1:D10). If omitted, gets all comments.'),
  responseFormat: responseFormatSchema,
});

export const addCommentSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  cellAddress: cellAddressSchema,
  text: z.string().describe('Comment text to add'),
  author: z.string().optional().describe('Comment author name'),
  createBackup: z.boolean().default(false),
});

// ============================================================
// Named Ranges
// ============================================================

export const listNamedRangesSchema = z.object({
  filePath: filePathSchema,
  responseFormat: responseFormatSchema,
});

export const createNamedRangeSchema = z.object({
  filePath: filePathSchema,
  name: z.string().describe('Name for the range (e.g., SalesData)'),
  sheetName: sheetNameSchema,
  range: rangeSchema.describe('Cell range to name (e.g., A1:D10)'),
  createBackup: z.boolean().default(false),
});

export const deleteNamedRangeSchema = z.object({
  filePath: filePathSchema,
  name: z.string().describe('Name of the named range to delete'),
  createBackup: z.boolean().default(false),
});

// ============================================================
// Sheet Protection
// ============================================================

export const setSheetProtectionSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  protect: z.boolean().describe('True to protect, false to unprotect'),
  password: z.string().optional().describe('Optional password for protection'),
  options: z.object({
    allowInsertRows: z.boolean().optional().default(false),
    allowInsertColumns: z.boolean().optional().default(false),
    allowDeleteRows: z.boolean().optional().default(false),
    allowDeleteColumns: z.boolean().optional().default(false),
    allowSort: z.boolean().optional().default(false),
    allowAutoFilter: z.boolean().optional().default(false),
    allowFormatCells: z.boolean().optional().default(false),
    allowFormatColumns: z.boolean().optional().default(false),
    allowFormatRows: z.boolean().optional().default(false),
  }).optional().describe('Protection options (what users are allowed to do)'),
  createBackup: z.boolean().default(false),
});

// ============================================================
// Data Validation (Set)
// ============================================================

export const setDataValidationSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.describe('Range to apply validation to (e.g., A1:A100)'),
  validationType: z.enum(['list', 'whole', 'decimal', 'date', 'textLength', 'custom']).describe('Type of validation'),
  operator: z.enum(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']).optional().describe('Comparison operator'),
  formula1: z.string().describe('First value/formula (for list: comma-separated values)'),
  formula2: z.string().optional().describe('Second value (for between/notBetween operators)'),
  showErrorMessage: z.boolean().optional().default(true),
  errorTitle: z.string().optional().default('Invalid Input'),
  errorMessage: z.string().optional().default('The value entered is not valid.'),
  createBackup: z.boolean().default(false),
});

// ============================================================
// Calculation Control (COM-only)
// ============================================================

export const triggerRecalculationSchema = z.object({
  filePath: filePathSchema,
  fullRecalc: z.boolean().optional().default(false).describe('If true, forces full recalculation of all formulas'),
});

export const getCalculationModeSchema = z.object({
  filePath: filePathSchema,
});

export const setCalculationModeSchema = z.object({
  filePath: filePathSchema,
  mode: z.enum(['automatic', 'manual', 'semiautomatic']).describe('Calculation mode to set'),
});

// ============================================================
// Screenshot (COM-only)
// ============================================================

export const captureScreenshotSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  range: rangeSchema.optional().describe('Optional range to capture. If omitted, captures the used range.'),
  outputPath: z.string().describe('File path to save the screenshot PNG'),
});

// ============================================================
// VBA Macros (COM-only)
// ============================================================

export const runVbaMacroSchema = z.object({
  filePath: filePathSchema,
  macroName: z.string().describe('Name of the macro to run (e.g., Sheet1.MyMacro or MyModule.MyMacro)'),
  args: z.array(z.any()).optional().default([]).describe('Arguments to pass to the macro'),
});

export const getVbaCodeSchema = z.object({
  filePath: filePathSchema,
  moduleName: z.string().describe('VBA module name (e.g., Module1, Sheet1, ThisWorkbook)'),
});

export const setVbaCodeSchema = z.object({
  filePath: filePathSchema,
  moduleName: z.string().describe('VBA module name (e.g., Module1)'),
  code: z.string().describe('VBA code to set in the module'),
  createModule: z.boolean().optional().default(false).describe('If true, creates a new module if it does not exist'),
  appendMode: z.boolean().optional().default(false).describe('If true, appends the code to existing module content instead of replacing it'),
});

// ============================================================
// Diagnosis (Connection & Accessibility)
// ============================================================

export const diagnoseConnectionSchema = z.object({
  filePath: filePathSchema.optional().describe('Optional path to a specific Excel file to check'),
  responseFormat: responseFormatSchema,
});

// ============================================================
// Power Query (COM-only)
// ============================================================

export const listPowerQueriesSchema = z.object({
  filePath: filePathSchema,
  responseFormat: responseFormatSchema,
});

export const checkVbaTrustSchema = z.object({});

export const enableVbaTrustSchema = z.object({
  enable: z.boolean().describe('True to enable VBA trust access, false to disable'),
});

export const runPowerQuerySchema = z.object({
  filePath: filePathSchema,
  queryName: z.string().describe('Name for the query'),
  formula: z.string().describe('M language formula for the query'),
  refreshOnly: z.boolean().optional().default(false).describe('If true, only refreshes an existing query instead of creating a new one'),
});

// ============================================================
// Batch Format (COM-only)
// ============================================================

export const batchFormatSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  operations: z.array(z.object({
    range: z.string().describe('Cell or range (e.g., "A1", "A1:D10")'),
    merge: z.boolean().optional().describe('Merge cells in range'),
    unmerge: z.boolean().optional().describe('Unmerge cells in range first'),
    value: z.union([z.string(), z.number()]).optional().describe('Set cell value'),
    fontName: z.string().optional().describe('Font name (e.g., "Segoe UI", "Calibri")'),
    fontSize: z.number().optional().describe('Font size in points'),
    fontBold: z.boolean().optional().describe('Bold text'),
    fontItalic: z.boolean().optional().describe('Italic text'),
    fontColor: z.string().optional().describe('Font color as hex (e.g., "#FFFFFF" for white)'),
    fillColor: z.string().optional().describe('Background fill color as hex (e.g., "#1E3247" for dark blue)'),
    horizontalAlignment: z.enum(['left', 'center', 'right']).optional().describe('Horizontal text alignment'),
    verticalAlignment: z.enum(['top', 'center', 'bottom']).optional().describe('Vertical text alignment'),
    numberFormat: z.string().optional().describe('Number format (e.g., "#,##0", "$#,##0.00", "0%")'),
    columnWidth: z.number().optional().describe('Set width for all columns in this range'),
    rowHeight: z.number().optional().describe('Set height for all rows in this range'),
    borderStyle: z.enum(['thin', 'medium', 'thick', 'none']).optional().describe('Border style for all edges'),
    borderColor: z.string().optional().describe('Border color as hex (e.g., "#000000")'),
    wrapText: z.boolean().optional().describe('Enable text wrapping'),
    autoFit: z.boolean().optional().describe('Auto-fit column widths to content'),
  })).describe('Array of formatting operations to apply in order'),
});

// ============================================================
// Display Options (COM-only)
// ============================================================

export const setDisplayOptionsSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema.optional().describe('Sheet to activate before applying options. If omitted, applies to active sheet.'),
  showGridlines: z.boolean().optional().describe('Show or hide gridlines. Hide for clean dashboard look.'),
  showRowColumnHeaders: z.boolean().optional().describe('Show or hide row numbers and column letters.'),
  zoomLevel: z.number().min(10).max(400).optional().describe('Zoom percentage (10-400). 85-100 is typical for dashboards.'),
  freezePaneCell: z.string().optional().describe('Cell address to freeze panes at (e.g., "A3" freezes rows 1-2). Set to "" to unfreeze.'),
  tabColor: z.string().optional().describe('Sheet tab color as hex (e.g., "#1E3247"). Set to "" to clear.'),
});

// ============================================================
// Shapes (COM-only)
// ============================================================

export const addShapeSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  shapeType: z.enum(['rectangle', 'roundedRectangle', 'oval']).describe('Shape type. Use roundedRectangle for dashboard cards.'),
  left: z.number().describe('Left position in points from worksheet edge'),
  top: z.number().describe('Top position in points from worksheet edge'),
  width: z.number().describe('Width in points'),
  height: z.number().describe('Height in points'),
  name: z.string().optional().describe('Name for the shape (for later reference)'),
  fill: z.object({
    color: z.string().optional().describe('Solid fill color as hex (e.g., "#1E3247")'),
    gradient: z.object({
      color1: z.string().describe('Start color hex'),
      color2: z.string().describe('End color hex'),
      direction: z.enum(['horizontal', 'vertical', 'diagonalDown', 'diagonalUp']).optional().default('vertical'),
    }).optional().describe('Two-color gradient fill (overrides solid color)'),
    transparency: z.number().min(0).max(1).optional().describe('Fill transparency 0-1 (0=opaque, 1=invisible)'),
  }).optional(),
  line: z.object({
    color: z.string().optional().describe('Border color hex'),
    weight: z.number().optional().describe('Border weight in points'),
    visible: z.boolean().optional().describe('Show or hide border. Set false for borderless cards.'),
  }).optional(),
  shadow: z.object({
    visible: z.boolean().optional().default(true),
    color: z.string().optional().default('#000000').describe('Shadow color hex'),
    offsetX: z.number().optional().default(3).describe('Horizontal shadow offset in points'),
    offsetY: z.number().optional().default(3).describe('Vertical shadow offset in points'),
    blur: z.number().optional().default(8).describe('Shadow blur radius in points'),
    transparency: z.number().min(0).max(1).optional().default(0.7).describe('Shadow transparency 0-1'),
  }).optional(),
  text: z.object({
    value: z.string().describe('Text content'),
    fontName: z.string().optional().default('Segoe UI'),
    fontSize: z.number().optional().default(12),
    fontBold: z.boolean().optional().default(false),
    fontColor: z.string().optional().default('#FFFFFF').describe('Text color hex'),
    horizontalAlignment: z.enum(['left', 'center', 'right']).optional().default('center'),
    verticalAlignment: z.enum(['top', 'middle', 'bottom']).optional().default('middle'),
    autoSize: z.enum(['none', 'shrinkToFit', 'shapeToFitText']).optional().default('none').describe('Auto-size text: shrinkToFit shrinks text to fit shape, shapeToFitText grows shape to fit text'),
  }).optional().describe('Text inside the shape'),
});

// ============================================================
// Chart Styling (COM-only)
// ============================================================

export const styleChartSchema = z.object({
  filePath: filePathSchema,
  sheetName: sheetNameSchema,
  chartIndex: z.number().optional().describe('Chart index (1-based) on the sheet'),
  chartName: z.string().optional().describe('Chart name (alternative to chartIndex)'),
  series: z.array(z.object({
    index: z.number().describe('Series index (1-based)'),
    color: z.string().optional().describe('Fill/line color hex (e.g., "#00BCD4")'),
    lineWeight: z.number().optional().describe('Line weight in points (line/scatter charts)'),
    markerStyle: z.enum(['circle', 'square', 'diamond', 'triangle', 'none']).optional(),
    markerSize: z.number().optional(),
    dataLabels: z.object({
      show: z.boolean(),
      numberFormat: z.string().optional().describe('e.g., "$#,##0", "0%"'),
      fontSize: z.number().optional(),
      fontColor: z.string().optional(),
      position: z.enum(['above', 'below', 'left', 'right', 'center', 'outsideEnd', 'insideEnd', 'insideBase']).optional(),
      hideBelow: z.number().optional().describe('Hide individual data labels where the absolute point value is below this threshold (e.g., 0.05 to hide labels on segments smaller than 5% in a percentage chart). Useful for stacked bar/column charts to prevent overlapping labels on tiny segments.'),
    }).optional(),
  })).optional(),
  axes: z.object({
    category: z.object({
      visible: z.boolean().optional(),
      numberFormat: z.string().optional(),
      fontSize: z.number().optional(),
      fontColor: z.string().optional(),
      labelRotation: z.number().optional(),
    }).optional(),
    value: z.object({
      visible: z.boolean().optional(),
      numberFormat: z.string().optional(),
      fontSize: z.number().optional(),
      fontColor: z.string().optional(),
      min: z.number().optional(),
      max: z.number().optional(),
      gridlines: z.boolean().optional(),
    }).optional(),
  }).optional(),
  chartArea: z.object({
    fillColor: z.string().optional(),
    borderVisible: z.boolean().optional(),
  }).optional(),
  plotArea: z.object({
    fillColor: z.string().optional(),
  }).optional(),
  legend: z.object({
    visible: z.boolean(),
    position: z.enum(['top', 'bottom', 'left', 'right']).optional(),
    fontSize: z.number().optional(),
    fontColor: z.string().optional(),
  }).optional(),
  title: z.object({
    text: z.string().optional(),
    visible: z.boolean().optional(),
    fontSize: z.number().optional(),
    fontColor: z.string().optional(),
  }).optional(),
  width: z.number().optional().describe('Chart width in points'),
  height: z.number().optional().describe('Chart height in points'),
});
