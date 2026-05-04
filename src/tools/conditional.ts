import { loadWorkbook, getSheet, saveWorkbook, parseRange } from './helpers.js';
import { isExcelRunningLive, isFileOpenInExcelLive, applyConditionalFormatLive, saveFileLive } from './excel-live.js';

interface RangeRect {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

function rangesOverlap(a: RangeRect, b: RangeRect): boolean {
  return !(a.endRow < b.startRow || a.startRow > b.endRow || a.endCol < b.startCol || a.startCol > b.endCol);
}

function tryParseMergeRect(s: string): RangeRect | null {
  const pair = s.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  const lettersToNum = (letters: string): number => {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  };
  if (pair) {
    return {
      startCol: lettersToNum(pair[1]),
      startRow: parseInt(pair[2]),
      endCol: lettersToNum(pair[3]),
      endRow: parseInt(pair[4]),
    };
  }
  const single = s.match(/^([A-Z]+)(\d+)$/);
  if (single) {
    const c = lettersToNum(single[1]);
    const r = parseInt(single[2]);
    return { startCol: c, startRow: r, endCol: c, endRow: r };
  }
  return null;
}

/**
 * Inspect the workbook on disk for merged ranges on the named sheet that
 * overlap the target CF range. Throws CF_RANGE_OVERLAPS_MERGED if any do.
 * Used as a precondition both in live (COM) and ExcelJS code paths so the
 * silent-failure mode is closed off everywhere.
 */
async function ensureNoMergedOverlap(
  filePath: string,
  sheetName: string,
  targetRect: RangeRect,
  rangeLabel: string
): Promise<void> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const merges: string[] = (sheet.model as any)?.merges || [];
  const overlapping: string[] = [];
  for (const m of merges) {
    const rect = tryParseMergeRect(m);
    if (!rect) continue;
    if (rangesOverlap(rect, targetRect)) overlapping.push(m);
  }
  if (overlapping.length > 0) {
    throw Object.assign(
      new Error(
        `Cannot apply conditional formatting to range ${rangeLabel} — it overlaps merged cell(s): ${overlapping.join(', ')}. Unmerge first via excel_unmerge_cells, or apply CF to a non-overlapping range.`
      ),
      { code: 'CF_RANGE_OVERLAPS_MERGED' }
    );
  }
}

export async function applyConditionalFormat(
  filePath: string,
  sheetName: string,
  range: string,
  ruleType: 'cellValue' | 'colorScale' | 'dataBar' | 'topBottom',
  condition?: {
    operator?: 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'notEqual' | 'containsText';
    value?: any;
    value2?: any;
  },
  style?: {
    font?: {
      color?: string;
      bold?: boolean;
    };
    fill?: {
      type: 'pattern';
      pattern: 'solid' | 'darkVertical' | 'darkHorizontal' | 'darkGrid';
      fgColor?: string;
    };
  },
  colorScale?: {
    minColor?: string;
    midColor?: string;
    maxColor?: string;
  },
  createBackup: boolean = false
): Promise<string> {
  // Parse range for validation
  const { startRow, startCol, endRow, endCol } = parseRange(range);
  const targetRect: RangeRect = { startRow, startCol, endRow, endCol };

  // Pre-flight: refuse to silently no-op when the target range overlaps merged
  // cells. ExcelJS (and even some COM paths) silently produce zero CF rules
  // in that case. Surface a structured error instead.
  await ensureNoMergedOverlap(filePath, sheetName, targetRect, range);

  // Check if Excel is running and file is open — use COM if so
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    await applyConditionalFormatLive(filePath, sheetName, range, ruleType, condition, style as any, colorScale);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Conditional formatting applied to ${range} via COM`,
      range,
      ruleType,
      method: 'live',
      note: 'Native Excel conditional formatting applied. Visible immediately in Excel.',
    }, null, 2);
  }

  // ExcelJS fallback
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  if (ruleType === 'cellValue' && condition && style) {
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const cellValue = cell.value;

        let shouldApplyStyle = false;

        switch (condition.operator) {
          case 'greaterThan':
            shouldApplyStyle = Number(cellValue) > Number(condition.value);
            break;
          case 'lessThan':
            shouldApplyStyle = Number(cellValue) < Number(condition.value);
            break;
          case 'equal':
            shouldApplyStyle = cellValue === condition.value;
            break;
          case 'notEqual':
            shouldApplyStyle = cellValue !== condition.value;
            break;
          case 'between':
            const numValue = Number(cellValue);
            shouldApplyStyle = numValue >= Number(condition.value) && numValue <= Number(condition.value2);
            break;
          case 'containsText':
            shouldApplyStyle = String(cellValue).includes(String(condition.value));
            break;
        }

        if (shouldApplyStyle) {
          if (style.font) {
            cell.font = {
              ...cell.font,
              color: style.font.color ? { argb: style.font.color } : cell.font?.color,
              bold: style.font.bold !== undefined ? style.font.bold : cell.font?.bold,
            };
          }

          if (style.fill) {
            cell.fill = {
              type: 'pattern',
              pattern: style.fill.pattern,
              fgColor: style.fill.fgColor ? { argb: style.fill.fgColor } : undefined,
            };
          }
        }
      }
    }
  } else if (ruleType === 'colorScale') {
    const minColor = colorScale?.minColor || 'FFFF0000';
    const maxColor = colorScale?.maxColor || 'FF00FF00';
    const midColor = colorScale?.midColor;

    const values: number[] = [];
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const value = Number(cell.value);
        if (!isNaN(value)) {
          values.push(value);
        }
      }
    }

    const minValue = Math.min(...values);
    const maxValue = Math.max(...values);
    const range_span = maxValue - minValue;

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = sheet.getRow(row).getCell(col);
        const value = Number(cell.value);

        if (!isNaN(value)) {
          const percentage = range_span === 0 ? 0.5 : (value - minValue) / range_span;

          let color: string;
          if (midColor && percentage < 0.5) {
            color = minColor;
          } else if (midColor && percentage >= 0.5) {
            color = maxColor;
          } else {
            color = percentage < 0.5 ? minColor : maxColor;
          }

          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color },
          };
        }
      }
    }
  }

  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({
    success: true,
    message: `Conditional formatting applied to ${range}`,
    range,
    ruleType,
    method: 'exceljs',
    note: 'Simplified ExcelJS implementation. Open file in Excel for native conditional formatting.',
  }, null, 2);
}
