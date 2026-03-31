import { loadWorkbook, saveWorkbook } from './helpers.js';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  listNamedRangesLive,
  createNamedRangeLive,
  deleteNamedRangeLive,
  saveFileLive,
} from './excel-live.js';
import type { ResponseFormat } from '../types.js';

export async function listNamedRanges(
  filePath: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    const raw = await listNamedRangesLive(filePath);
    const names = raw ? JSON.parse(raw) : [];

    if (responseFormat === 'markdown') {
      if (!names.length) return '# Named Ranges\n\nNo named ranges found.';
      let md = '# Named Ranges\n\n| Name | Refers To | Visible |\n|------|-----------|--------|\n';
      for (const n of names) {
        md += `| ${n.Name} | ${n.RefersTo} | ${n.Visible} |\n`;
      }
      return md;
    }

    return JSON.stringify({ namedRanges: names, method: 'live' }, null, 2);
  } else {
    const workbook = await loadWorkbook(filePath);
    const names: Array<{ name: string; refersTo: string }> = [];

    if ((workbook as any).definedNames) {
      const dn = (workbook as any).definedNames;
      if (dn.model) {
        for (const entry of dn.model) {
          names.push({ name: entry.name, refersTo: entry.ranges?.join(',') || '' });
        }
      }
    }

    if (responseFormat === 'markdown') {
      if (!names.length) return '# Named Ranges\n\nNo named ranges found.';
      let md = '# Named Ranges\n\n| Name | Refers To |\n|------|-----------|\n';
      for (const n of names) {
        md += `| ${n.name} | ${n.refersTo} |\n`;
      }
      return md;
    }

    return JSON.stringify({ namedRanges: names, method: 'exceljs' }, null, 2);
  }
}

export async function createNamedRange(
  filePath: string,
  name: string,
  sheetName: string,
  range: string,
  createBackup: boolean = false
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    await createNamedRangeLive(filePath, name, sheetName, range);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Named range "${name}" created for ${sheetName}!${range}`,
      name,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    const workbook = await loadWorkbook(filePath);

    // ExcelJS defined names
    if ((workbook as any).definedNames) {
      (workbook as any).definedNames.add(`'${sheetName}'!${range}`, name);
    }

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Named range "${name}" created for ${sheetName}!${range}`,
      name,
      method: 'exceljs',
    }, null, 2);
  }
}

export async function deleteNamedRange(
  filePath: string,
  name: string,
  createBackup: boolean = false
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    await deleteNamedRangeLive(filePath, name);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Named range "${name}" deleted`,
      name,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    const workbook = await loadWorkbook(filePath);

    if ((workbook as any).definedNames) {
      (workbook as any).definedNames.remove(name);
    }

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Named range "${name}" deleted`,
      name,
      method: 'exceljs',
    }, null, 2);
  }
}
