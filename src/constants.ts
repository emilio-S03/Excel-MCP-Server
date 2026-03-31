export const TOOL_ANNOTATIONS = {
  READ_ONLY: { readOnlyHint: true },
  DESTRUCTIVE: { destructiveHint: true },
  IDEMPOTENT: { idempotentHint: true },
};

export const ERROR_MESSAGES = {
  FILE_NOT_FOUND: 'Excel file not found',
  SHEET_NOT_FOUND: 'Sheet not found',
  INVALID_RANGE: 'Invalid cell range',
  INVALID_CELL: 'Invalid cell address',
  WRITE_ERROR: 'Error writing to Excel file',
  READ_ERROR: 'Error reading Excel file',
  INVALID_FORMAT: 'Invalid format specification',
  FILE_LOCKED: 'File is currently open in another application (like Excel). Please close the file and try again.',
  EXCEL_NOT_RUNNING: 'This feature requires Excel to be running with the file open. Please open the file in Excel first.',
  VBA_TRUST_CENTER: 'VBA access denied. Enable "Trust access to the VBA project object model" in Excel Trust Center settings.',
  POWER_QUERY_WARNING: 'Power Query may access external data sources. Ensure you trust the data source before running.',
};

export const DEFAULT_OPTIONS = {
  RESPONSE_FORMAT: 'json' as const,
  CREATE_BACKUP: false,
  MAX_ROWS_DISPLAY: 100,
};
