/**
 * Structured error codes attached to thrown errors via `(err as any).code`.
 * The dispatcher in src/index.ts surfaces these in the JSON tool response as
 * `errorCode`, so callers can switch on them without parsing message text.
 *
 * Add a code here whenever you `Object.assign(new Error(...), { code: '...' })`
 * in a tool, so the canonical list stays in one place.
 */
export const ERROR_CODES = {
  PATH_OUTSIDE_ALLOWED: 'PATH_OUTSIDE_ALLOWED',
  PLATFORM_UNSUPPORTED: 'PLATFORM_UNSUPPORTED',
  EXCEL_NOT_RUNNING: 'EXCEL_NOT_RUNNING',
  VBA_TRUST_DENIED: 'VBA_TRUST_DENIED',
  COM_TIMEOUT: 'COM_TIMEOUT',
  COM_UNREACHABLE: 'COM_UNREACHABLE',
  CF_RANGE_OVERLAPS_MERGED: 'CF_RANGE_OVERLAPS_MERGED',
  INVALID_CELL: 'INVALID_CELL',
  INVALID_RANGE: 'INVALID_RANGE',
  SHEET_NOT_FOUND: 'SHEET_NOT_FOUND',
  FILE_NOT_FOUND: 'FILE_NOT_FOUND',
  WRITE_ERROR: 'WRITE_ERROR',
  READ_ERROR: 'READ_ERROR',
} as const;

export type ErrorCode = typeof ERROR_CODES[keyof typeof ERROR_CODES];
