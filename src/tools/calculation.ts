import { isExcelRunningLive, isFileOpenInExcelLive } from './excel-live.js';
import {
  triggerRecalculationLive,
  getCalculationModeLive,
  setCalculationModeLive,
} from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

async function ensureFileOpenInExcel(filePath: string): Promise<void> {
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
}

export async function triggerRecalculation(
  filePath: string,
  fullRecalc: boolean = false
): Promise<string> {
  await ensureFileOpenInExcel(filePath);
  await triggerRecalculationLive(filePath, fullRecalc);

  return JSON.stringify({
    success: true,
    message: fullRecalc ? 'Full recalculation triggered' : 'Recalculation triggered',
    method: 'live',
  }, null, 2);
}

export async function getCalculationMode(
  filePath: string
): Promise<string> {
  await ensureFileOpenInExcel(filePath);
  const mode = await getCalculationModeLive(filePath);

  return JSON.stringify({
    mode,
    method: 'live',
  }, null, 2);
}

export async function setCalculationMode(
  filePath: string,
  mode: string
): Promise<string> {
  await ensureFileOpenInExcel(filePath);
  await setCalculationModeLive(filePath, mode);

  return JSON.stringify({
    success: true,
    message: `Calculation mode set to "${mode}"`,
    mode,
    method: 'live',
  }, null, 2);
}
