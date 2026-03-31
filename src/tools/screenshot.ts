import { isExcelRunningLive, isFileOpenInExcelLive, captureScreenshotLive } from './excel-live.js';
import { ensureFilePathAllowed } from './helpers.js';
import { ERROR_MESSAGES } from '../constants.js';

export async function captureScreenshot(
  filePath: string,
  sheetName: string,
  outputPath: string,
  range?: string
): Promise<string> {
  // Validate output path is in an allowed directory
  ensureFilePathAllowed(outputPath);

  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  await captureScreenshotLive(filePath, sheetName, outputPath, range);

  return JSON.stringify({
    success: true,
    message: `Screenshot saved to ${outputPath}`,
    outputPath,
    range: range || 'used range',
    method: 'live',
  }, null, 2);
}
