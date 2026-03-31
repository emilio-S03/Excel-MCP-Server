#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  InitializeRequestSchema,
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { ZodError } from 'zod';

// Import tool implementations
import { readWorkbook, readSheet, readRange, getCell, getFormula } from './tools/read.js';
import { writeWorkbook, updateCell, writeRange, addRow, setFormula } from './tools/write.js';
import { formatCell, setColumnWidth, setRowHeight, mergeCells, batchFormat } from './tools/format.js';
import { createSheet, deleteSheet, renameSheet, duplicateSheet, setSheetProtection } from './tools/sheets.js';
import { deleteRows, deleteColumns, copyRange } from './tools/operations.js';
import { searchValue, filterRows } from './tools/analysis.js';
import { createChart, styleChart } from './tools/charts.js';
import { createPivotTable } from './tools/pivots.js';
import { createTable } from './tools/tables.js';
import { validateFormulaSyntax, validateExcelRange, getDataValidationInfo, setDataValidation } from './tools/validation.js';
import { insertRows, insertColumns, unmergeCells, getMergedCells } from './tools/advanced.js';
import { applyConditionalFormat } from './tools/conditional.js';
import { getComments, addComment } from './tools/comments.js';
import { listNamedRanges, createNamedRange, deleteNamedRange } from './tools/named-ranges.js';
import { triggerRecalculation, getCalculationMode, setCalculationMode } from './tools/calculation.js';
import { captureScreenshot } from './tools/screenshot.js';
import { runVbaMacro, getVbaCode, setVbaCode, checkVbaTrust, enableVbaTrust } from './tools/vba.js';
import { diagnoseConnection } from './tools/diagnose.js';
import { listPowerQueries, runPowerQuery } from './tools/power-query.js';
import { setDisplayOptions } from './tools/display.js';
import { addShape } from './tools/shapes.js';

import { TOOL_ANNOTATIONS } from './constants.js';
import * as schemas from './schemas/index.js';
import { setAllowedDirectories } from './tools/helpers.js';

// User configuration storage
interface UserConfig {
  createBackupByDefault?: boolean;
  defaultResponseFormat?: 'json' | 'markdown';
  allowedDirectories?: string[];
}

let userConfig: UserConfig = {
  createBackupByDefault: false,
  defaultResponseFormat: 'json',
  allowedDirectories: [],
};

// Schema mapping for validation
const toolSchemas: Record<string, any> = {
  excel_read_workbook: schemas.readWorkbookSchema,
  excel_read_sheet: schemas.readSheetSchema,
  excel_read_range: schemas.readRangeSchema,
  excel_get_cell: schemas.getCellSchema,
  excel_get_formula: schemas.getFormulaSchema,
  excel_write_workbook: schemas.writeWorkbookSchema,
  excel_update_cell: schemas.updateCellSchema,
  excel_write_range: schemas.writeRangeSchema,
  excel_add_row: schemas.addRowSchema,
  excel_set_formula: schemas.setFormulaSchema,
  excel_format_cell: schemas.formatCellSchema,
  excel_set_column_width: schemas.setColumnWidthSchema,
  excel_set_row_height: schemas.setRowHeightSchema,
  excel_merge_cells: schemas.mergeCellsSchema,
  excel_create_sheet: schemas.createSheetSchema,
  excel_delete_sheet: schemas.deleteSheetSchema,
  excel_rename_sheet: schemas.renameSheetSchema,
  excel_duplicate_sheet: schemas.duplicateSheetSchema,
  excel_delete_rows: schemas.deleteRowsSchema,
  excel_delete_columns: schemas.deleteColumnsSchema,
  excel_copy_range: schemas.copyRangeSchema,
  excel_search_value: schemas.searchValueSchema,
  excel_filter_rows: schemas.filterRowsSchema,
  excel_create_chart: schemas.createChartSchema,
  excel_style_chart: schemas.styleChartSchema,
  excel_create_pivot_table: schemas.createPivotTableSchema,
  excel_create_table: schemas.createTableSchema,
  excel_validate_formula_syntax: schemas.validateFormulaSyntaxSchema,
  excel_validate_range: schemas.validateExcelRangeSchema,
  excel_get_data_validation_info: schemas.getDataValidationInfoSchema,
  excel_insert_rows: schemas.insertRowsSchema,
  excel_insert_columns: schemas.insertColumnsSchema,
  excel_unmerge_cells: schemas.unmergeCellsSchema,
  excel_get_merged_cells: schemas.getMergedCellsSchema,
  excel_apply_conditional_format: schemas.applyConditionalFormatSchema,
  // New tools
  excel_get_comments: schemas.getCommentsSchema,
  excel_add_comment: schemas.addCommentSchema,
  excel_list_named_ranges: schemas.listNamedRangesSchema,
  excel_create_named_range: schemas.createNamedRangeSchema,
  excel_delete_named_range: schemas.deleteNamedRangeSchema,
  excel_set_sheet_protection: schemas.setSheetProtectionSchema,
  excel_set_data_validation: schemas.setDataValidationSchema,
  excel_trigger_recalculation: schemas.triggerRecalculationSchema,
  excel_get_calculation_mode: schemas.getCalculationModeSchema,
  excel_set_calculation_mode: schemas.setCalculationModeSchema,
  excel_capture_screenshot: schemas.captureScreenshotSchema,
  excel_run_vba_macro: schemas.runVbaMacroSchema,
  excel_get_vba_code: schemas.getVbaCodeSchema,
  excel_set_vba_code: schemas.setVbaCodeSchema,
  excel_list_power_queries: schemas.listPowerQueriesSchema,
  excel_run_power_query: schemas.runPowerQuerySchema,
  excel_check_vba_trust: schemas.checkVbaTrustSchema,
  excel_enable_vba_trust: schemas.enableVbaTrustSchema,
  excel_diagnose_connection: schemas.diagnoseConnectionSchema,
  excel_batch_format: schemas.batchFormatSchema,
  excel_set_display_options: schemas.setDisplayOptionsSchema,
  excel_add_shape: schemas.addShapeSchema,
};

// Create server instance
const server = new Server(
  {
    name: 'excel-mcp-server',
    version: '2.0.0',
  },
  {
    capabilities: {
      tools: {
        listChanged: true,
      },
    },
  }
);

// Handle initialization
server.setRequestHandler(InitializeRequestSchema, async () => {
  return {
    protocolVersion: '2024-11-05',
    capabilities: {
      tools: {},
    },
    serverInfo: {
      name: 'excel-mcp-server',
      version: '2.0.0',
    },
  };
});

// List all available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      // READ OPERATIONS
      {
        name: 'excel_read_workbook',
        description: 'List all sheets and metadata of an Excel workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_read_sheet',
        description: 'Read complete data from a sheet (with optional range)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Optional range (e.g., A1:D10)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_read_range',
        description: 'Read a specific range of cells',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to read (e.g., A1:D10)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_get_cell',
        description: 'Read value from a specific cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'cellAddress'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_get_formula',
        description: 'Read the formula from a specific cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'cellAddress'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // WRITE OPERATIONS
      {
        name: 'excel_write_workbook',
        description: 'Create a new Excel file with data',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path for the new Excel file' },
            sheetName: { type: 'string', description: 'Name for the sheet', default: 'Sheet1' },
            data: { type: 'array', description: '2D array of data to write' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'data'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_update_cell',
        description: 'Update value of a specific cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            value: { description: 'Value to write' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'value'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_write_range',
        description: 'Write multiple cells simultaneously',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to write (e.g., A1:D10)' },
            data: { type: 'array', description: '2D array of data to write' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'data'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_add_row',
        description: 'Add a row at the end of the sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            data: { type: 'array', description: 'Array of values for the new row' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'data'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_set_formula',
        description: 'Set or modify a formula in a cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            formula: { type: 'string', description: 'Excel formula (without = sign)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'formula'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // FORMAT OPERATIONS
      {
        name: 'excel_format_cell',
        description: 'Change cell formatting (color, font, borders, alignment). For formatting 3+ cells, prefer excel_batch_format instead (faster, one call). Use Segoe UI font, not default Calibri. Professional color palette: dark navy #1E3247, medium navy #283D52, cyan accent #00BCD4, body text #424242, borders #E0E0E0.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            format: { type: 'object', description: 'Format options' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'format'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_set_column_width',
        description: 'Adjust width of a column. ALWAYS set widths explicitly — never leave defaults. Guidelines: names/titles=20-25, currency=12-14, percentages=8-10, dates=12, scores/counts=8-10, short codes=8-10, spacer columns=2-3. For multiple columns, prefer excel_batch_format with columnWidth property.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            column: { description: 'Column letter (A) or number (1)' },
            width: { type: 'number', description: 'Width in Excel units' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'column', 'width'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_set_row_height',
        description: 'Adjust height of a row',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            row: { type: 'number', description: 'Row number (1-based)' },
            height: { type: 'number', description: 'Height in points' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'row', 'height'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_merge_cells',
        description: 'Merge cells in a range',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to merge (e.g., A1:D1)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // SHEET MANAGEMENT
      {
        name: 'excel_create_sheet',
        description: 'Create a new sheet in the workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name for the new sheet' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_delete_sheet',
        description: 'Delete a sheet from the workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_rename_sheet',
        description: 'Rename a sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            oldName: { type: 'string', description: 'Current sheet name' },
            newName: { type: 'string', description: 'New sheet name' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'oldName', 'newName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_duplicate_sheet',
        description: 'Duplicate a complete sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sourceSheetName: { type: 'string', description: 'Name of sheet to duplicate' },
            newSheetName: { type: 'string', description: 'Name for duplicated sheet' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sourceSheetName', 'newSheetName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // OPERATIONS
      {
        name: 'excel_delete_rows',
        description: 'Delete specific rows',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startRow: { type: 'number', description: 'Starting row number (1-based)' },
            count: { type: 'number', description: 'Number of rows to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startRow', 'count'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_delete_columns',
        description: 'Delete specific columns',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startColumn: { description: 'Starting column (letter or number)' },
            count: { type: 'number', description: 'Number of columns to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startColumn', 'count'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_copy_range',
        description: 'Copy range to another location',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sourceSheetName: { type: 'string', description: 'Source sheet name' },
            sourceRange: { type: 'string', description: 'Source range (e.g., A1:D10)' },
            targetSheetName: { type: 'string', description: 'Target sheet name' },
            targetCell: { type: 'string', description: 'Top-left cell of destination' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sourceSheetName', 'sourceRange', 'targetSheetName', 'targetCell'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // ANALYSIS
      {
        name: 'excel_search_value',
        description: 'Search for a value in sheet/range',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            searchValue: { description: 'Value to search for' },
            range: { type: 'string', description: 'Optional range to search within' },
            caseSensitive: { type: 'boolean', default: false },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'searchValue'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_filter_rows',
        description: 'Filter rows by condition',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            column: { description: 'Column to filter by' },
            condition: { type: 'string', enum: ['equals', 'contains', 'greater_than', 'less_than', 'not_empty'] },
            value: { description: 'Value to compare against' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'column', 'condition'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // CHARTS
      {
        name: 'excel_create_chart',
        description: 'Create a chart (line, bar, column, pie, scatter, area). DESIGN RULES: No 3D effects, no gradients, no chart borders. Use flat fills only. Primary color: #1E3247 (dark navy). Secondary: #00BCD4 (cyan). Tertiary: #009688 (teal). Chart title: 11pt bold #424242. Axis labels: 9pt #424242. Gridlines: light gray #E0E0E0, thin. Remove vertical gridlines for bar charts. Legend: 9pt, bottom or right position. Pie charts: max 6-8 slices, group rest as "Other", show percentages not values. Bar charts: sort by value (largest first) unless categorical order matters. Line charts: 2pt line weight, circle markers 4pt.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            chartType: { type: 'string', enum: ['line', 'bar', 'column', 'pie', 'scatter', 'area'] },
            dataRange: { type: 'string', description: 'Range of data (e.g., A1:D10)' },
            dataSheetName: { type: 'string', description: 'Sheet containing the data range (if different from sheetName where chart is placed)' },
            position: { type: 'string', description: 'Position for chart (e.g., F2)' },
            title: { type: 'string', description: 'Chart title' },
            showLegend: { type: 'boolean', default: true },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'chartType', 'dataRange', 'position'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_style_chart',
        description: 'Style an existing chart — series colors, data labels, axis formatting, chart/plot area fills, legend, title, and size. Replaces VBA macro styling with a single native call.\n\nSERIES COLORS: Set fill+line color per series index. Use #1E3247 (navy), #00BCD4 (cyan), #009688 (teal), #FF7043 (coral).\nDATA LABELS: Show with number format (e.g., "$#,##0"), position (above/center/outsideEnd), font size+color.\nAXES: Hide value axis for clean look, format category labels (rotation, font size), set min/max scale, toggle gridlines.\nCHART/PLOT AREA: Set background fill colors. Use transparent/white for charts layered on dashboard cards.\nLEGEND: Position top/bottom/left/right, set font size+color, or hide entirely.\nTITLE: Set text, font size+color, or hide.\n\nIdentify chart by chartIndex (1-based) or chartName.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            sheetName: { type: 'string', description: 'Sheet containing the chart' },
            chartIndex: { type: 'number', description: 'Chart index (1-based) on the sheet' },
            chartName: { type: 'string', description: 'Chart name (alternative to chartIndex)' },
            series: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  index: { type: 'number', description: 'Series index (1-based)' },
                  color: { type: 'string', description: 'Fill/line color hex (e.g., "#00BCD4")' },
                  lineWeight: { type: 'number', description: 'Line weight in points' },
                  markerStyle: { type: 'string', enum: ['circle', 'square', 'diamond', 'triangle', 'none'] },
                  markerSize: { type: 'number' },
                  dataLabels: {
                    type: 'object',
                    properties: {
                      show: { type: 'boolean' },
                      numberFormat: { type: 'string', description: 'e.g., "$#,##0", "0%"' },
                      fontSize: { type: 'number' },
                      fontColor: { type: 'string' },
                      position: { type: 'string', enum: ['above', 'below', 'left', 'right', 'center', 'outsideEnd', 'insideEnd', 'insideBase'] },
                    },
                    required: ['show'],
                  },
                },
                required: ['index'],
              },
            },
            axes: {
              type: 'object',
              properties: {
                category: {
                  type: 'object',
                  properties: {
                    visible: { type: 'boolean' },
                    numberFormat: { type: 'string' },
                    fontSize: { type: 'number' },
                    fontColor: { type: 'string' },
                    labelRotation: { type: 'number' },
                  },
                },
                value: {
                  type: 'object',
                  properties: {
                    visible: { type: 'boolean' },
                    numberFormat: { type: 'string' },
                    fontSize: { type: 'number' },
                    fontColor: { type: 'string' },
                    min: { type: 'number' },
                    max: { type: 'number' },
                    gridlines: { type: 'boolean' },
                  },
                },
              },
            },
            chartArea: {
              type: 'object',
              properties: {
                fillColor: { type: 'string' },
                borderVisible: { type: 'boolean' },
              },
            },
            plotArea: {
              type: 'object',
              properties: {
                fillColor: { type: 'string' },
              },
            },
            legend: {
              type: 'object',
              properties: {
                visible: { type: 'boolean' },
                position: { type: 'string', enum: ['top', 'bottom', 'left', 'right'] },
                fontSize: { type: 'number' },
                fontColor: { type: 'string' },
              },
              required: ['visible'],
            },
            title: {
              type: 'object',
              properties: {
                text: { type: 'string' },
                visible: { type: 'boolean' },
                fontSize: { type: 'number' },
                fontColor: { type: 'string' },
              },
            },
            width: { type: 'number', description: 'Chart width in points' },
            height: { type: 'number', description: 'Chart height in points' },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // PIVOT TABLES
      {
        name: 'excel_create_pivot_table',
        description: 'Create a pivot table for data analysis. After creation, use excel_batch_format to style it: header row with #3A5068 fill and white bold text, alternating data rows white/#FAFAFA, total row with #E8F5E9 fill and bold #2E7D32 text, thin borders #E0E0E0.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sourceSheetName: { type: 'string', description: 'Source sheet name' },
            sourceRange: { type: 'string', description: 'Source data range' },
            targetSheetName: { type: 'string', description: 'Target sheet for pivot table' },
            targetCell: { type: 'string', description: 'Target cell (e.g., A1)' },
            rows: { type: 'array', items: { type: 'string' }, description: 'Row fields' },
            columns: { type: 'array', items: { type: 'string' }, description: 'Column fields' },
            values: { type: 'array', description: 'Value fields with aggregation' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sourceSheetName', 'sourceRange', 'targetSheetName', 'targetCell', 'rows', 'values'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // TABLES
      {
        name: 'excel_create_table',
        description: 'Convert a range to an Excel table with formatting. Use TableStyleMedium2 or TableStyleDark1 for professional look. After creating the table, use excel_batch_format to apply custom header colors (#3A5068 fill, white bold text) and alternating row colors.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to convert (e.g., A1:D10)' },
            tableName: { type: 'string', description: 'Name for the table' },
            tableStyle: { type: 'string', default: 'TableStyleMedium2' },
            showFirstColumn: { type: 'boolean', default: false },
            showLastColumn: { type: 'boolean', default: false },
            showRowStripes: { type: 'boolean', default: true },
            showColumnStripes: { type: 'boolean', default: false },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'tableName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // VALIDATION
      {
        name: 'excel_validate_formula_syntax',
        description: 'Validate Excel formula syntax without applying it',
        inputSchema: {
          type: 'object',
          properties: {
            formula: { type: 'string', description: 'Formula to validate (without = sign)' },
          },
          required: ['formula'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_validate_range',
        description: 'Validate if a range string is valid',
        inputSchema: {
          type: 'object',
          properties: {
            range: { type: 'string', description: 'Range to validate (e.g., A1:D10)' },
          },
          required: ['range'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_get_data_validation_info',
        description: 'Get data validation rules for a cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName', 'cellAddress'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // ADVANCED OPERATIONS
      {
        name: 'excel_insert_rows',
        description: 'Insert rows at a specific position',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startRow: { type: 'number', description: 'Row number to insert at (1-based)' },
            count: { type: 'number', description: 'Number of rows to insert' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startRow', 'count'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_insert_columns',
        description: 'Insert columns at a specific position',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            startColumn: { description: 'Column to insert at (letter or number)' },
            count: { type: 'number', description: 'Number of columns to insert' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'startColumn', 'count'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_unmerge_cells',
        description: 'Unmerge previously merged cells',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to unmerge (e.g., A1:D1)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_get_merged_cells',
        description: 'List all merged cell ranges in a sheet',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // CONDITIONAL FORMATTING
      {
        name: 'excel_apply_conditional_format',
        description: 'Apply conditional formatting to a range. Use semantic colors: positive values = green fill #E8F5E9 with #2E7D32 text, negative = red fill #FFEBEE with #C62828 text, warnings = orange fill #FFF3E0 with #E65100 text. For color scales, use #C62828 (low/bad) to #2E7D32 (high/good). Data bars: use #1E3247 (dark navy).',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to format (e.g., A1:D10)' },
            ruleType: { type: 'string', enum: ['cellValue', 'colorScale', 'dataBar', 'topBottom'] },
            condition: { type: 'object', description: 'Condition for cellValue type' },
            style: { type: 'object', description: 'Style to apply' },
            colorScale: { type: 'object', description: 'Color scale settings' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'ruleType'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // COMMENTS
      {
        name: 'excel_get_comments',
        description: 'Get all comments/notes from a sheet or range',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Optional range to get comments from' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath', 'sheetName'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_add_comment',
        description: 'Add a comment/note to a cell',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            cellAddress: { type: 'string', description: 'Cell address (e.g., A1)' },
            text: { type: 'string', description: 'Comment text' },
            author: { type: 'string', description: 'Comment author name' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'cellAddress', 'text'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // NAMED RANGES
      {
        name: 'excel_list_named_ranges',
        description: 'List all named ranges in a workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_create_named_range',
        description: 'Create a named range in a workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            name: { type: 'string', description: 'Name for the range (e.g., SalesData)' },
            sheetName: { type: 'string', description: 'Sheet containing the range' },
            range: { type: 'string', description: 'Cell range (e.g., A1:D10)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'name', 'sheetName', 'range'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_delete_named_range',
        description: 'Delete a named range from a workbook',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            name: { type: 'string', description: 'Name of the named range to delete' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'name'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // SHEET PROTECTION
      {
        name: 'excel_set_sheet_protection',
        description: 'Protect or unprotect a sheet with optional password and permissions',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            protect: { type: 'boolean', description: 'True to protect, false to unprotect' },
            password: { type: 'string', description: 'Optional protection password' },
            options: { type: 'object', description: 'Protection options (what users are allowed to do)' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'protect'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // DATA VALIDATION
      {
        name: 'excel_set_data_validation',
        description: 'Set data validation rules on a range (dropdowns, number limits, etc.)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Range to validate (e.g., A1:A100)' },
            validationType: { type: 'string', enum: ['list', 'whole', 'decimal', 'date', 'textLength', 'custom'] },
            operator: { type: 'string', enum: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] },
            formula1: { type: 'string', description: 'First value/formula (for list: comma-separated values)' },
            formula2: { type: 'string', description: 'Second value (for between/notBetween)' },
            showErrorMessage: { type: 'boolean', default: true },
            errorTitle: { type: 'string', default: 'Invalid Input' },
            errorMessage: { type: 'string', default: 'The value entered is not valid.' },
            createBackup: { type: 'boolean', default: false },
          },
          required: ['filePath', 'sheetName', 'range', 'validationType', 'formula1'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // CALCULATION CONTROL (COM-only)
      {
        name: 'excel_trigger_recalculation',
        description: 'Trigger recalculation of all formulas (requires Excel running)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            fullRecalc: { type: 'boolean', default: false, description: 'Force full recalculation' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.IDEMPOTENT,
      },
      {
        name: 'excel_get_calculation_mode',
        description: 'Get the current calculation mode (automatic/manual/semiautomatic, requires Excel running)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_set_calculation_mode',
        description: 'Set calculation mode to automatic, manual, or semiautomatic (requires Excel running)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            mode: { type: 'string', enum: ['automatic', 'manual', 'semiautomatic'] },
          },
          required: ['filePath', 'mode'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // SCREENSHOT (COM-only)
      {
        name: 'excel_capture_screenshot',
        description: 'Capture a screenshot of a sheet or range as PNG (requires Excel running)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            sheetName: { type: 'string', description: 'Name of the sheet' },
            range: { type: 'string', description: 'Optional range to capture' },
            outputPath: { type: 'string', description: 'File path to save the PNG' },
          },
          required: ['filePath', 'sheetName', 'outputPath'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // VBA MACROS (COM-only)
      {
        name: 'excel_run_vba_macro',
        description: 'Run a VBA macro in the workbook (requires Excel running with VBA trust enabled)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            macroName: { type: 'string', description: 'Macro name (e.g., Sheet1.MyMacro)' },
            args: { type: 'array', description: 'Arguments to pass to the macro' },
          },
          required: ['filePath', 'macroName'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
      {
        name: 'excel_get_vba_code',
        description: 'Read VBA code from a module (requires Excel running with VBA trust enabled)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            moduleName: { type: 'string', description: 'VBA module name (e.g., Module1, Sheet1)' },
          },
          required: ['filePath', 'moduleName'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_set_vba_code',
        description: 'Write VBA code to a module (requires Excel running with VBA trust enabled)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            moduleName: { type: 'string', description: 'VBA module name (e.g., Module1)' },
            code: { type: 'string', description: 'VBA code to write' },
            createModule: { type: 'boolean', default: false, description: 'Create module if it does not exist' },
            appendMode: { type: 'boolean', default: false, description: 'Append code to existing module instead of replacing all content' },
          },
          required: ['filePath', 'moduleName', 'code'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // VBA TRUST SETTINGS
      {
        name: 'excel_check_vba_trust',
        description: 'Check if VBA trust access is enabled in Windows registry (needed for VBA tools)',
        inputSchema: {
          type: 'object',
          properties: {},
          required: [],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_enable_vba_trust',
        description: 'Enable or disable VBA trust access in Windows registry. Requires Excel restart. This changes a security setting.',
        inputSchema: {
          type: 'object',
          properties: {
            enable: { type: 'boolean', description: 'True to enable VBA trust, false to disable' },
          },
          required: ['enable'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // DIAGNOSIS
      {
        name: 'excel_diagnose_connection',
        description: 'Diagnose Excel COM connection issues. Runs a series of checks: Excel process running, COM reachable, Excel responsive (no modal dialogs), file open, VBA trust enabled. Use this when other tools fail with COM errors.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Optional path to an Excel file to check if it is open and VBA-accessible' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: [],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },

      // POWER QUERY (COM-only)
      {
        name: 'excel_list_power_queries',
        description: 'List all Power Query queries in the workbook (requires Excel running)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            responseFormat: { type: 'string', enum: ['json', 'markdown'], default: 'json' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.READ_ONLY,
      },
      {
        name: 'excel_run_power_query',
        description: 'Create or refresh a Power Query (M language) query (requires Excel running)',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            queryName: { type: 'string', description: 'Name for the query' },
            formula: { type: 'string', description: 'M language formula' },
            refreshOnly: { type: 'boolean', default: false, description: 'Only refresh existing query' },
          },
          required: ['filePath', 'queryName', 'formula'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // BATCH FORMAT (COM-only)
      {
        name: 'excel_batch_format',
        description: 'Apply multiple formatting operations in one call (merge, colors, fonts, widths, borders, values). Much faster and more reliable than individual format calls or VBA macros. Requires Excel running with the file open. PREFERRED tool for dashboard design, report polish, and bulk formatting — use this instead of VBA macros.\n\nDESIGN GUIDE — Follow these rules for professional output:\n• Layout: Row 1 = title bar (merged, dark fill #1E3247, white 14pt bold, height 40). Row 2 = accent line (fill #00BCD4, height 4). Row 3+ = section headers (fill #283D52, white 11pt bold). Then column headers (fill #3A5068, white 10pt bold), then data rows.\n• Font: Use "Segoe UI" for everything. Body text = 10pt #424242. Headers = bold white. Never use default Calibri 11.\n• Colors: Dark navy #1E3247 (titles), medium navy #283D52 (sections), slate #3A5068 (col headers), cyan accent #00BCD4 (dividers/highlights), white #FFFFFF (text on dark), off-white #FAFAFA (alternating rows), light gray #F5F5F5 (card backgrounds), border gray #E0E0E0.\n• Semantic: Input cells = yellow fill #FFF9C4, calculated = blue fill #E3F2FD, totals = green fill #E8F5E9 with bold #2E7D32 text, errors = red #C62828, positive = green #1B5E20.\n• Data rows: Alternating white/#FAFAFA, thin borders #E0E0E0, 10pt, height 20. Numbers right-aligned. Text left-aligned.\n• Column widths: ALWAYS set explicitly. Names=20-25, currency=12-14, percentages=8-10, dates=12, scores=8-10. Never leave defaults.\n• Always use unmerge:true before merge:true to prevent errors on re-formatting.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            sheetName: { type: 'string', description: 'Name of the sheet to format' },
            operations: {
              type: 'array',
              description: 'Array of formatting operations applied in order. Each operation targets a range and can set multiple properties at once.',
              items: {
                type: 'object',
                properties: {
                  range: { type: 'string', description: 'Cell or range (e.g., "A1", "A1:I1")' },
                  merge: { type: 'boolean', description: 'Merge cells in range' },
                  unmerge: { type: 'boolean', description: 'Unmerge first (use when re-formatting previously merged cells)' },
                  value: { description: 'Set cell value (string or number)' },
                  fontName: { type: 'string', description: 'Font name (e.g., "Segoe UI", "Calibri")' },
                  fontSize: { type: 'number', description: 'Font size in points' },
                  fontBold: { type: 'boolean', description: 'Bold text' },
                  fontItalic: { type: 'boolean', description: 'Italic text' },
                  fontColor: { type: 'string', description: 'Font color hex (e.g., "#FFFFFF")' },
                  fillColor: { type: 'string', description: 'Background color hex (e.g., "#1E3247")' },
                  horizontalAlignment: { type: 'string', enum: ['left', 'center', 'right'], description: 'Horizontal alignment' },
                  verticalAlignment: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Vertical alignment' },
                  numberFormat: { type: 'string', description: 'Number format (e.g., "#,##0", "$#,##0.00")' },
                  columnWidth: { type: 'number', description: 'Column width for all columns in range' },
                  rowHeight: { type: 'number', description: 'Row height for all rows in range' },
                  borderStyle: { type: 'string', enum: ['thin', 'medium', 'thick', 'none'], description: 'Border style' },
                  borderColor: { type: 'string', description: 'Border color hex' },
                  wrapText: { type: 'boolean', description: 'Enable text wrapping' },
                  autoFit: { type: 'boolean', description: 'Auto-fit column widths to content' },
                },
                required: ['range'],
              },
            },
          },
          required: ['filePath', 'sheetName', 'operations'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // DISPLAY OPTIONS (COM-only)
      {
        name: 'excel_set_display_options',
        description: 'Control worksheet display: hide/show gridlines, row/column headers, set zoom level, freeze panes, set tab color. CRITICAL FOR DASHBOARDS: Always hide gridlines (showGridlines: false) and hide headers (showRowColumnHeaders: false) for any dashboard or designed layout. This transforms a spreadsheet into a clean canvas. Set zoom to 85-100% for dashboards.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            sheetName: { type: 'string', description: 'Sheet to apply options to (optional, uses active sheet)' },
            showGridlines: { type: 'boolean', description: 'Show/hide gridlines. HIDE for dashboards.' },
            showRowColumnHeaders: { type: 'boolean', description: 'Show/hide row numbers and column letters' },
            zoomLevel: { type: 'number', description: 'Zoom percentage (10-400)' },
            freezePaneCell: { type: 'string', description: 'Cell to freeze panes at (e.g., "A3"). Empty string to unfreeze.' },
            tabColor: { type: 'string', description: 'Sheet tab color hex. Empty string to clear.' },
          },
          required: ['filePath'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },

      // SHAPES (COM-only)
      {
        name: 'excel_add_shape',
        description: 'Add a shape to a worksheet — rectangles, rounded rectangles, ovals with fills, gradients, shadows, borders, and text. This is the key tool for DASHBOARD CARD LAYOUTS — the technique that makes Excel look like a designed application instead of a spreadsheet.\n\nDASHBOARD CARD PATTERN: Use roundedRectangle shapes as card containers. Place them over the cell grid to create visual sections. Each card typically has: dark fill (#1E3247 or #283D52), no border (line.visible: false), subtle shadow (blur: 8, transparency: 0.7), and white bold text for the card title.\n\nPOSITIONING: Values are in points (1 inch = 72 points). A typical column is ~48-64 points wide. Row height ~15 points. For a card starting at column B row 3: left≈64, top≈45.\n\nCOMBINE WITH excel_set_display_options (hide gridlines) + excel_batch_format (format data cells beneath/beside shapes) for full dashboard designs.',
        inputSchema: {
          type: 'object',
          properties: {
            filePath: { type: 'string', description: 'Path to the Excel file (must be open in Excel)' },
            sheetName: { type: 'string', description: 'Sheet to add the shape to' },
            shapeType: { type: 'string', enum: ['rectangle', 'roundedRectangle', 'oval'], description: 'Shape type. roundedRectangle for cards.' },
            left: { type: 'number', description: 'Left position in points' },
            top: { type: 'number', description: 'Top position in points' },
            width: { type: 'number', description: 'Width in points' },
            height: { type: 'number', description: 'Height in points' },
            name: { type: 'string', description: 'Shape name for later reference' },
            fill: {
              type: 'object',
              description: 'Fill options: solid color, gradient, or transparency',
              properties: {
                color: { type: 'string', description: 'Solid fill hex color' },
                gradient: {
                  type: 'object',
                  properties: {
                    color1: { type: 'string', description: 'Start color hex' },
                    color2: { type: 'string', description: 'End color hex' },
                    direction: { type: 'string', enum: ['horizontal', 'vertical', 'diagonalDown', 'diagonalUp'] },
                  },
                  required: ['color1', 'color2'],
                },
                transparency: { type: 'number', description: '0-1 (0=opaque)' },
              },
            },
            line: {
              type: 'object',
              properties: {
                color: { type: 'string' },
                weight: { type: 'number' },
                visible: { type: 'boolean', description: 'false for borderless cards' },
              },
            },
            shadow: {
              type: 'object',
              properties: {
                visible: { type: 'boolean' },
                color: { type: 'string' },
                offsetX: { type: 'number' },
                offsetY: { type: 'number' },
                blur: { type: 'number' },
                transparency: { type: 'number' },
              },
            },
            text: {
              type: 'object',
              properties: {
                value: { type: 'string' },
                fontName: { type: 'string' },
                fontSize: { type: 'number' },
                fontBold: { type: 'boolean' },
                fontColor: { type: 'string' },
                horizontalAlignment: { type: 'string', enum: ['left', 'center', 'right'] },
                verticalAlignment: { type: 'string', enum: ['top', 'middle', 'bottom'] },
                autoSize: { type: 'string', enum: ['none', 'shrinkToFit', 'shapeToFitText'], description: 'shrinkToFit auto-shrinks text to fit shape, shapeToFitText grows shape to fit text' },
              },
              required: ['value'],
            },
          },
          required: ['filePath', 'sheetName', 'shapeType', 'left', 'top', 'width', 'height'],
        },
        annotations: TOOL_ANNOTATIONS.DESTRUCTIVE,
      },
    ],
  };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const { name, arguments: args } = request.params;

    if (!args) {
      throw new McpError(ErrorCode.InvalidParams, 'Missing arguments');
    }

    // Validate arguments with Zod schema
    const schema = toolSchemas[name];
    if (!schema) {
      throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }

    let validatedArgs: any;
    try {
      validatedArgs = schema.parse(args);
    } catch (error) {
      if (error instanceof ZodError) {
        const issues = error.errors.map(e => `${e.path.join('.')}: ${e.message}`).join(', ');
        throw new McpError(ErrorCode.InvalidParams, `Invalid arguments: ${issues}`);
      }
      throw error;
    }

    // Apply user config defaults
    if (validatedArgs.createBackup === undefined && userConfig.createBackupByDefault !== undefined) {
      validatedArgs.createBackup = userConfig.createBackupByDefault;
    }
    if (validatedArgs.responseFormat === undefined && userConfig.defaultResponseFormat !== undefined) {
      validatedArgs.responseFormat = userConfig.defaultResponseFormat;
    }

    let result: string;

    switch (name) {
      // Read operations
      case 'excel_read_workbook':
        result = await readWorkbook(validatedArgs.filePath, validatedArgs.responseFormat);
        break;
      case 'excel_read_sheet':
        result = await readSheet(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.responseFormat);
        break;
      case 'excel_read_range':
        result = await readRange(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.responseFormat);
        break;
      case 'excel_get_cell':
        result = await getCell(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.responseFormat);
        break;
      case 'excel_get_formula':
        result = await getFormula(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.responseFormat);
        break;

      // Write operations
      case 'excel_write_workbook':
        result = await writeWorkbook(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.data, validatedArgs.createBackup);
        break;
      case 'excel_update_cell':
        result = await updateCell(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.value, validatedArgs.createBackup);
        break;
      case 'excel_write_range':
        result = await writeRange(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.data, validatedArgs.createBackup);
        break;
      case 'excel_add_row':
        result = await addRow(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.data, validatedArgs.createBackup);
        break;
      case 'excel_set_formula':
        result = await setFormula(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.formula, validatedArgs.createBackup);
        break;

      // Format operations
      case 'excel_format_cell':
        result = await formatCell(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.cellAddress, validatedArgs.format, validatedArgs.createBackup);
        break;
      case 'excel_set_column_width':
        result = await setColumnWidth(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.column, validatedArgs.width, validatedArgs.createBackup);
        break;
      case 'excel_set_row_height':
        result = await setRowHeight(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.row, validatedArgs.height, validatedArgs.createBackup);
        break;
      case 'excel_merge_cells':
        result = await mergeCells(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.range, validatedArgs.createBackup);
        break;

      // Sheet management
      case 'excel_create_sheet':
        result = await createSheet(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.createBackup);
        break;
      case 'excel_delete_sheet':
        result = await deleteSheet(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.createBackup);
        break;
      case 'excel_rename_sheet':
        result = await renameSheet(validatedArgs.filePath, validatedArgs.oldName, validatedArgs.newName, validatedArgs.createBackup);
        break;
      case 'excel_duplicate_sheet':
        result = await duplicateSheet(validatedArgs.filePath, validatedArgs.sourceSheetName, validatedArgs.newSheetName, validatedArgs.createBackup);
        break;

      // Operations
      case 'excel_delete_rows':
        result = await deleteRows(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.startRow, validatedArgs.count, validatedArgs.createBackup);
        break;
      case 'excel_delete_columns':
        result = await deleteColumns(validatedArgs.filePath, validatedArgs.sheetName, validatedArgs.startColumn, validatedArgs.count, validatedArgs.createBackup);
        break;
      case 'excel_copy_range':
        result = await copyRange(
          validatedArgs.filePath,
          validatedArgs.sourceSheetName,
          validatedArgs.sourceRange,
          validatedArgs.targetSheetName,
          validatedArgs.targetCell,
          validatedArgs.createBackup
        );
        break;

      // Analysis
      case 'excel_search_value':
        result = await searchValue(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.searchValue,
          validatedArgs.range,
          validatedArgs.caseSensitive,
          validatedArgs.responseFormat
        );
        break;
      case 'excel_filter_rows':
        result = await filterRows(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.column,
          validatedArgs.condition,
          validatedArgs.value,
          validatedArgs.responseFormat
        );
        break;

      // Charts
      case 'excel_create_chart':
        result = await createChart(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.chartType,
          validatedArgs.dataRange,
          validatedArgs.position,
          validatedArgs.title,
          validatedArgs.showLegend,
          validatedArgs.createBackup,
          validatedArgs.dataSheetName
        );
        break;

      case 'excel_style_chart':
        result = await styleChart(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.chartIndex,
          validatedArgs.chartName,
          {
            series: validatedArgs.series,
            axes: validatedArgs.axes,
            chartArea: validatedArgs.chartArea,
            plotArea: validatedArgs.plotArea,
            legend: validatedArgs.legend,
            title: validatedArgs.title,
            width: validatedArgs.width,
            height: validatedArgs.height,
          }
        );
        break;

      // Pivot tables
      case 'excel_create_pivot_table':
        result = await createPivotTable(
          validatedArgs.filePath,
          validatedArgs.sourceSheetName,
          validatedArgs.sourceRange,
          validatedArgs.targetSheetName,
          validatedArgs.targetCell,
          validatedArgs.rows,
          validatedArgs.columns,
          validatedArgs.values,
          validatedArgs.createBackup
        );
        break;

      // Tables
      case 'excel_create_table':
        result = await createTable(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.tableName,
          validatedArgs.tableStyle,
          validatedArgs.showFirstColumn,
          validatedArgs.showLastColumn,
          validatedArgs.showRowStripes,
          validatedArgs.showColumnStripes,
          validatedArgs.createBackup
        );
        break;

      // Validation
      case 'excel_validate_formula_syntax':
        result = await validateFormulaSyntax(validatedArgs.formula);
        break;

      case 'excel_validate_range':
        result = await validateExcelRange(validatedArgs.range);
        break;

      case 'excel_get_data_validation_info':
        result = await getDataValidationInfo(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.cellAddress,
          validatedArgs.responseFormat
        );
        break;

      // Advanced operations
      case 'excel_insert_rows':
        result = await insertRows(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.startRow,
          validatedArgs.count,
          validatedArgs.createBackup
        );
        break;

      case 'excel_insert_columns':
        result = await insertColumns(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.startColumn,
          validatedArgs.count,
          validatedArgs.createBackup
        );
        break;

      case 'excel_unmerge_cells':
        result = await unmergeCells(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.createBackup
        );
        break;

      case 'excel_get_merged_cells':
        result = await getMergedCells(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.responseFormat
        );
        break;

      // Conditional formatting
      case 'excel_apply_conditional_format':
        result = await applyConditionalFormat(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.ruleType,
          validatedArgs.condition,
          validatedArgs.style,
          validatedArgs.colorScale,
          validatedArgs.createBackup
        );
        break;

      // Comments
      case 'excel_get_comments':
        result = await getComments(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.responseFormat
        );
        break;

      case 'excel_add_comment':
        result = await addComment(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.cellAddress,
          validatedArgs.text,
          validatedArgs.author,
          validatedArgs.createBackup
        );
        break;

      // Named Ranges
      case 'excel_list_named_ranges':
        result = await listNamedRanges(
          validatedArgs.filePath,
          validatedArgs.responseFormat
        );
        break;

      case 'excel_create_named_range':
        result = await createNamedRange(
          validatedArgs.filePath,
          validatedArgs.name,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.createBackup
        );
        break;

      case 'excel_delete_named_range':
        result = await deleteNamedRange(
          validatedArgs.filePath,
          validatedArgs.name,
          validatedArgs.createBackup
        );
        break;

      // Sheet Protection
      case 'excel_set_sheet_protection':
        result = await setSheetProtection(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.protect,
          validatedArgs.password,
          validatedArgs.options,
          validatedArgs.createBackup
        );
        break;

      // Data Validation
      case 'excel_set_data_validation':
        result = await setDataValidation(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.range,
          validatedArgs.validationType,
          validatedArgs.formula1,
          validatedArgs.operator,
          validatedArgs.formula2,
          validatedArgs.showErrorMessage,
          validatedArgs.errorTitle,
          validatedArgs.errorMessage,
          validatedArgs.createBackup
        );
        break;

      // Calculation Control (COM-only)
      case 'excel_trigger_recalculation':
        result = await triggerRecalculation(
          validatedArgs.filePath,
          validatedArgs.fullRecalc
        );
        break;

      case 'excel_get_calculation_mode':
        result = await getCalculationMode(validatedArgs.filePath);
        break;

      case 'excel_set_calculation_mode':
        result = await setCalculationMode(
          validatedArgs.filePath,
          validatedArgs.mode
        );
        break;

      // Screenshot (COM-only)
      case 'excel_capture_screenshot':
        result = await captureScreenshot(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.outputPath,
          validatedArgs.range
        );
        break;

      // VBA Macros (COM-only)
      case 'excel_run_vba_macro':
        result = await runVbaMacro(
          validatedArgs.filePath,
          validatedArgs.macroName,
          validatedArgs.args
        );
        break;

      case 'excel_get_vba_code':
        result = await getVbaCode(
          validatedArgs.filePath,
          validatedArgs.moduleName
        );
        break;

      case 'excel_set_vba_code':
        result = await setVbaCode(
          validatedArgs.filePath,
          validatedArgs.moduleName,
          validatedArgs.code,
          validatedArgs.createModule,
          validatedArgs.appendMode
        );
        break;

      // VBA Trust Settings
      case 'excel_check_vba_trust':
        result = await checkVbaTrust();
        break;

      case 'excel_enable_vba_trust':
        result = await enableVbaTrust(validatedArgs.enable);
        break;

      // Diagnosis
      case 'excel_diagnose_connection':
        result = await diagnoseConnection(validatedArgs.filePath, validatedArgs.responseFormat);
        break;

      // Power Query (COM-only)
      case 'excel_list_power_queries':
        result = await listPowerQueries(
          validatedArgs.filePath,
          validatedArgs.responseFormat
        );
        break;

      case 'excel_run_power_query':
        result = await runPowerQuery(
          validatedArgs.filePath,
          validatedArgs.queryName,
          validatedArgs.formula,
          validatedArgs.refreshOnly
        );
        break;

      // Batch Format (COM-only)
      case 'excel_batch_format':
        result = await batchFormat(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.operations
        );
        break;

      // Display Options (COM-only)
      case 'excel_set_display_options':
        result = await setDisplayOptions(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          validatedArgs.showGridlines,
          validatedArgs.showRowColumnHeaders,
          validatedArgs.zoomLevel,
          validatedArgs.freezePaneCell,
          validatedArgs.tabColor
        );
        break;

      // Shapes (COM-only)
      case 'excel_add_shape':
        result = await addShape(
          validatedArgs.filePath,
          validatedArgs.sheetName,
          {
            shapeType: validatedArgs.shapeType,
            left: validatedArgs.left,
            top: validatedArgs.top,
            width: validatedArgs.width,
            height: validatedArgs.height,
            name: validatedArgs.name,
            fill: validatedArgs.fill,
            line: validatedArgs.line,
            shadow: validatedArgs.shadow,
            text: validatedArgs.text,
          }
        );
        break;

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }

    return {
      content: [
        {
          type: 'text',
          text: result,
        },
      ],
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ error: errorMessage }, null, 2),
        },
      ],
      isError: true,
    };
  }
});

// Handle configuration updates via notifications
// Note: This will be called by Claude Desktop when user updates config
async function handleConfigUpdate(config: any) {
  try {
    if (config.createBackupByDefault !== undefined) {
      userConfig.createBackupByDefault = config.createBackupByDefault;
    }

    if (config.defaultResponseFormat !== undefined) {
      userConfig.defaultResponseFormat = config.defaultResponseFormat;
    }

    if (config.allowedDirectories !== undefined) {
      userConfig.allowedDirectories = Array.isArray(config.allowedDirectories)
        ? config.allowedDirectories
        : [];

      // Update allowed directories in helpers
      setAllowedDirectories(userConfig.allowedDirectories || []);
    }

    console.error('Configuration updated:', userConfig);
  } catch (error) {
    console.error('Error handling configuration:', error);
  }
}

// Set up notification handler
server.notification = async (notification: any) => {
  if (notification.method === 'notifications/configure') {
    await handleConfigUpdate(notification.params?.config || notification.params);
  }
};

// Handle EPIPE errors gracefully
process.stdout.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EPIPE') {
    // Ignore EPIPE errors - they happen when the client disconnects
    process.exit(0);
  }
});

process.on('uncaughtException', (err: Error & { code?: string }) => {
  if (err.code === 'EPIPE') {
    // Ignore EPIPE errors
    process.exit(0);
  }
  console.error('Uncaught exception:', err);
  process.exit(1);
});

// Start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Excel MCP Server running on stdio');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
