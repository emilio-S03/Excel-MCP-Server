import { loadWorkbook, getSheet, saveWorkbook } from './helpers.js';
import {
  isExcelRunningLive,
  isFileOpenInExcelLive,
  getCommentsLive,
  addCommentLive,
  saveFileLive,
} from './excel-live.js';
import type { ResponseFormat } from '../types.js';

export async function getComments(
  filePath: string,
  sheetName: string,
  range?: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    const raw = await getCommentsLive(filePath, sheetName, range);
    const comments = raw ? JSON.parse(raw) : [];

    if (responseFormat === 'markdown') {
      if (!comments.length) return '# Comments\n\nNo comments found.';
      let md = `# Comments in ${sheetName}\n\n`;
      for (const c of comments) {
        md += `- **${c.Address}** (${c.Author || 'Unknown'}): ${c.Text}\n`;
      }
      return md;
    }

    return JSON.stringify({ comments, method: 'live' }, null, 2);
  } else {
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);

    const comments: Array<{ address: string; author: string; text: string }> = [];

    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.note) {
          const note = typeof cell.note === 'string'
            ? { author: '', text: cell.note }
            : { author: (cell.note as any).author || '', text: (cell.note as any).text || String(cell.note) };
          comments.push({
            address: cell.address,
            author: note.author,
            text: note.text,
          });
        }
      });
    });

    if (responseFormat === 'markdown') {
      if (!comments.length) return '# Comments\n\nNo comments found.';
      let md = `# Comments in ${sheetName}\n\n`;
      for (const c of comments) {
        md += `- **${c.address}** (${c.author || 'Unknown'}): ${c.text}\n`;
      }
      return md;
    }

    return JSON.stringify({ comments, method: 'exceljs' }, null, 2);
  }
}

export async function addComment(
  filePath: string,
  sheetName: string,
  cellAddress: string,
  text: string,
  author?: string,
  createBackup: boolean = false
): Promise<string> {
  const excelRunning = await isExcelRunningLive();
  const fileOpen = excelRunning ? await isFileOpenInExcelLive(filePath) : false;

  if (fileOpen) {
    await addCommentLive(filePath, sheetName, cellAddress, text, author);
    await saveFileLive(filePath);

    return JSON.stringify({
      success: true,
      message: `Comment added to ${cellAddress}`,
      cellAddress,
      method: 'live',
      note: 'Changes visible immediately in Excel',
    }, null, 2);
  } else {
    const workbook = await loadWorkbook(filePath);
    const sheet = getSheet(workbook, sheetName);
    const cell = sheet.getCell(cellAddress);

    if (author) {
      cell.note = { texts: [{ text }], author } as any;
    } else {
      cell.note = text;
    }

    await saveWorkbook(workbook, filePath, createBackup);

    return JSON.stringify({
      success: true,
      message: `Comment added to ${cellAddress}`,
      cellAddress,
      method: 'exceljs',
    }, null, 2);
  }
}
