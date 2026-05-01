/**
 * v3.1 — page setup, print area, headers/footers.
 * File-mode read/write. PDF export is in pdf-export.ts (live mode).
 */
import { loadWorkbook, getSheet, saveWorkbook } from './helpers.js';

export async function getPageSetup(filePath: string, sheetName: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const ps: any = sheet.pageSetup ?? {};
  const hf: any = (sheet as any).headerFooter ?? {};
  return JSON.stringify({
    sheetName,
    orientation: ps.orientation ?? 'portrait',
    paperSize: ps.paperSize,
    fitToPage: ps.fitToPage,
    fitToWidth: ps.fitToWidth,
    fitToHeight: ps.fitToHeight,
    scale: ps.scale,
    horizontalCentered: ps.horizontalCentered,
    verticalCentered: ps.verticalCentered,
    printArea: ps.printArea,
    printTitlesRow: ps.printTitlesRow,
    printTitlesColumn: ps.printTitlesColumn,
    margins: ps.margins,
    blackAndWhite: ps.blackAndWhite,
    draft: ps.draft,
    pageOrder: ps.pageOrder,
    headerFooter: {
      oddHeader: hf.oddHeader,
      oddFooter: hf.oddFooter,
      evenHeader: hf.evenHeader,
      evenFooter: hf.evenFooter,
      differentFirst: hf.differentFirst,
      differentOddEven: hf.differentOddEven,
    },
  }, null, 2);
}

export async function setPageSetup(
  filePath: string,
  sheetName: string,
  config: {
    orientation?: 'portrait' | 'landscape';
    paperSize?: number;
    fitToPage?: boolean;
    fitToWidth?: number;
    fitToHeight?: number;
    scale?: number;
    horizontalCentered?: boolean;
    verticalCentered?: boolean;
    printArea?: string;
    printTitlesRow?: string;
    printTitlesColumn?: string;
    margins?: {
      left?: number;
      right?: number;
      top?: number;
      bottom?: number;
      header?: number;
      footer?: number;
    };
    headerFooter?: {
      oddHeader?: string;
      oddFooter?: string;
      evenHeader?: string;
      evenFooter?: string;
    };
  },
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const ps: any = sheet.pageSetup ?? (sheet.pageSetup = {} as any);

  if (config.orientation !== undefined) ps.orientation = config.orientation;
  if (config.paperSize !== undefined) ps.paperSize = config.paperSize;
  if (config.fitToPage !== undefined) ps.fitToPage = config.fitToPage;
  if (config.fitToWidth !== undefined) ps.fitToWidth = config.fitToWidth;
  if (config.fitToHeight !== undefined) ps.fitToHeight = config.fitToHeight;
  if (config.scale !== undefined) ps.scale = config.scale;
  if (config.horizontalCentered !== undefined) ps.horizontalCentered = config.horizontalCentered;
  if (config.verticalCentered !== undefined) ps.verticalCentered = config.verticalCentered;
  if (config.printArea !== undefined) ps.printArea = config.printArea;
  if (config.printTitlesRow !== undefined) ps.printTitlesRow = config.printTitlesRow;
  if (config.printTitlesColumn !== undefined) ps.printTitlesColumn = config.printTitlesColumn;
  if (config.margins) {
    ps.margins = { ...(ps.margins ?? {}), ...config.margins };
  }

  if (config.headerFooter) {
    const hf: any = (sheet as any).headerFooter ?? ((sheet as any).headerFooter = {} as any);
    if (config.headerFooter.oddHeader !== undefined) hf.oddHeader = config.headerFooter.oddHeader;
    if (config.headerFooter.oddFooter !== undefined) hf.oddFooter = config.headerFooter.oddFooter;
    if (config.headerFooter.evenHeader !== undefined) hf.evenHeader = config.headerFooter.evenHeader;
    if (config.headerFooter.evenFooter !== undefined) hf.evenFooter = config.headerFooter.evenFooter;
  }

  await saveWorkbook(workbook, filePath, createBackup);
  return JSON.stringify({ success: true, sheetName, applied: config }, null, 2);
}
