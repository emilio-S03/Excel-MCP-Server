import { isExcelRunningLive, isFileOpenInExcelLive } from './excel-live.js';
import { listPowerQueriesLive, runPowerQueryLive, saveFileLive } from './excel-live.js';
import { ERROR_MESSAGES } from '../constants.js';
import type { ResponseFormat } from '../types.js';

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

export async function listPowerQueries(
  filePath: string,
  responseFormat: ResponseFormat = 'json'
): Promise<string> {
  await ensureFileOpenInExcel(filePath);
  const raw = await listPowerQueriesLive(filePath);
  const queries = raw ? JSON.parse(raw) : [];

  if (responseFormat === 'markdown') {
    if (!queries.length) return '# Power Queries\n\nNo queries found.';
    let md = '# Power Queries\n\n';
    for (const q of queries) {
      md += `## ${q.Name}\n`;
      if (q.Description) md += `*${q.Description}*\n\n`;
      md += `\`\`\`\n${q.Formula}\n\`\`\`\n\n`;
    }
    return md;
  }

  return JSON.stringify({
    queries,
    method: 'live',
    note: ERROR_MESSAGES.POWER_QUERY_WARNING,
  }, null, 2);
}

export async function runPowerQuery(
  filePath: string,
  queryName: string,
  formula: string,
  refreshOnly: boolean = false
): Promise<string> {
  await ensureFileOpenInExcel(filePath);

  console.error(`[PowerQuery] WARNING: ${ERROR_MESSAGES.POWER_QUERY_WARNING}`);

  await runPowerQueryLive(filePath, queryName, formula, refreshOnly);
  await saveFileLive(filePath);

  return JSON.stringify({
    success: true,
    message: refreshOnly
      ? `Query "${queryName}" refreshed`
      : `Query "${queryName}" created/updated`,
    queryName,
    method: 'live',
    note: ERROR_MESSAGES.POWER_QUERY_WARNING,
  }, null, 2);
}
