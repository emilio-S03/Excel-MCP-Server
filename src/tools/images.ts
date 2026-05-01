import { promises as fs } from 'fs';
import { extname } from 'path';
import { loadWorkbook, getSheet, saveWorkbook, parseRange, columnLetterToNumber } from './helpers.js';

const SUPPORTED_EXT: Record<string, 'png' | 'jpeg' | 'gif'> = {
  '.png': 'png',
  '.jpg': 'jpeg',
  '.jpeg': 'jpeg',
  '.gif': 'gif',
};

function detectExtension(imagePath: string): 'png' | 'jpeg' | 'gif' {
  const ext = extname(imagePath).toLowerCase();
  const mapped = SUPPORTED_EXT[ext];
  if (!mapped) {
    throw new Error(
      `Unsupported image format: ${ext || '(no extension)'}. Supported: ${Object.keys(SUPPORTED_EXT).join(', ')}`
    );
  }
  return mapped;
}

export async function addImage(
  filePath: string,
  sheetName: string,
  imagePath: string,
  options: {
    cell?: string;
    range?: string;
    widthPx?: number;
    heightPx?: number;
    createBackup?: boolean;
  }
): Promise<string> {
  await fs.access(imagePath);

  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const buffer = await fs.readFile(imagePath);
  const extension = detectExtension(imagePath);
  const imageId = workbook.addImage({ buffer: buffer as any, extension });

  if (options.range) {
    const { startCol, startRow, endCol, endRow } = parseRange(options.range);
    sheet.addImage(imageId, {
      tl: { col: startCol - 1, row: startRow - 1 } as any,
      br: { col: endCol, row: endRow } as any,
      editAs: 'oneCell',
    });
  } else {
    const cell = options.cell ?? 'A1';
    const match = cell.match(/^([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Invalid cell address: ${cell}`);
    const colIdx = columnLetterToNumber(match[1]) - 1;
    const rowIdx = parseInt(match[2], 10) - 1;

    if (options.widthPx && options.heightPx) {
      sheet.addImage(imageId, {
        tl: { col: colIdx, row: rowIdx } as any,
        ext: { width: options.widthPx, height: options.heightPx },
        editAs: 'oneCell',
      });
    } else {
      sheet.addImage(imageId, `${cell}:${cell}` as any);
    }
  }

  await saveWorkbook(workbook, filePath, options.createBackup ?? false);

  return JSON.stringify(
    {
      success: true,
      message: `Image added to ${sheetName} at ${options.range ?? options.cell ?? 'A1'}`,
      imagePath,
      filePath,
      sheetName,
      placement: options.range ?? options.cell ?? 'A1',
    },
    null,
    2
  );
}
