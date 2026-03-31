import { isExcelRunningLive, isFileOpenInExcelLive, addShapeLive } from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';

export async function addShape(
  filePath: string,
  sheetName: string,
  config: {
    shapeType: string;
    left: number;
    top: number;
    width: number;
    height: number;
    name?: string;
    fill?: {
      color?: string;
      gradient?: { color1: string; color2: string; direction?: string };
      transparency?: number;
    };
    line?: { color?: string; weight?: number; visible?: boolean };
    shadow?: { visible?: boolean; color?: string; offsetX?: number; offsetY?: number; blur?: number; transparency?: number };
    text?: { value: string; fontName?: string; fontSize?: number; fontBold?: boolean; fontColor?: string; horizontalAlignment?: string; verticalAlignment?: string; autoSize?: string };
  }
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  if (!excelRunning) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }
  const fileOpen = await isFileOpenInExcelLive(filePath);
  if (!fileOpen) {
    throw new Error(ERROR_MESSAGES.EXCEL_NOT_RUNNING);
  }

  const shapeName = await addShapeLive(filePath, sheetName, config);

  return JSON.stringify({
    success: true,
    message: `Shape "${shapeName}" created on sheet "${sheetName}"`,
    shapeName: shapeName.trim(),
    shapeType: config.shapeType,
    position: { left: config.left, top: config.top, width: config.width, height: config.height },
    method: 'live',
    note: 'Shape visible immediately in Excel. Use the shape name to reference it later.',
  }, null, 2);
}
