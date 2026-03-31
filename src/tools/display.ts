import { isExcelRunningLive, isFileOpenInExcelLive, setDisplayOptionsLive } from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

export async function setDisplayOptions(
  filePath: string,
  sheetName?: string,
  showGridlines?: boolean,
  showRowColumnHeaders?: boolean,
  zoomLevel?: number,
  freezePaneCell?: string,
  tabColor?: string
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  await setDisplayOptionsLive(filePath, sheetName, showGridlines, showRowColumnHeaders, zoomLevel, freezePaneCell, tabColor);

  const changes: string[] = [];
  if (showGridlines !== undefined) changes.push(`gridlines ${showGridlines ? 'shown' : 'hidden'}`);
  if (showRowColumnHeaders !== undefined) changes.push(`headers ${showRowColumnHeaders ? 'shown' : 'hidden'}`);
  if (zoomLevel !== undefined) changes.push(`zoom set to ${zoomLevel}%`);
  if (freezePaneCell !== undefined) changes.push(freezePaneCell ? `panes frozen at ${freezePaneCell}` : 'panes unfrozen');
  if (tabColor !== undefined) changes.push(tabColor ? `tab color set` : 'tab color cleared');

  return JSON.stringify({
    success: true,
    message: `Display options updated: ${changes.join(', ')}`,
    method: 'live',
    note: 'Changes visible immediately in Excel.',
  }, null, 2);
}
