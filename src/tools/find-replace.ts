import {
  loadWorkbook,
  getSheet,
  saveWorkbook,
  cellValueToString,
  columnNumberToLetter,
} from './helpers.js';

interface FindReplaceMatch {
  sheet: string;
  cell: string;
  before: string;
  after: string;
}

export async function findReplace(
  filePath: string,
  pattern: string,
  replacement: string,
  options: {
    sheetName?: string;
    regex?: boolean;
    caseSensitive?: boolean;
    dryRun?: boolean;
    createBackup?: boolean;
  }
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const matches: FindReplaceMatch[] = [];

  const matcher = options.regex
    ? new RegExp(pattern, options.caseSensitive ? 'g' : 'gi')
    : null;

  const literal = options.caseSensitive ? pattern : pattern.toLowerCase();

  const targetSheets = options.sheetName ? [getSheet(workbook, options.sheetName)] : workbook.worksheets;

  for (const sheet of targetSheets) {
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const original = cellValueToString(cell.value);
        if (!original) return;

        let replaced: string | null = null;

        if (matcher) {
          if (matcher.test(original)) {
            matcher.lastIndex = 0;
            replaced = original.replace(matcher, replacement);
          }
        } else {
          const haystack = options.caseSensitive ? original : original.toLowerCase();
          if (haystack.includes(literal)) {
            if (options.caseSensitive) {
              replaced = original.split(pattern).join(replacement);
            } else {
              const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              replaced = original.replace(new RegExp(escaped, 'gi'), replacement);
            }
          }
        }

        if (replaced !== null && replaced !== original) {
          const cellAddr = `${columnNumberToLetter(colNumber)}${rowNumber}`;
          matches.push({
            sheet: sheet.name,
            cell: cellAddr,
            before: original,
            after: replaced,
          });
          if (!options.dryRun) {
            const cellValue = cell.value;
            if (cellValue && typeof cellValue === 'object' && 'formula' in cellValue) {
              cell.value = { ...cellValue, result: replaced } as any;
            } else {
              cell.value = replaced;
            }
          }
        }
      });
    });
  }

  if (!options.dryRun && matches.length > 0) {
    await saveWorkbook(workbook, filePath, options.createBackup ?? false);
  }

  return JSON.stringify(
    {
      success: true,
      filePath,
      pattern,
      replacement,
      mode: options.dryRun ? 'dryRun' : 'applied',
      regex: !!options.regex,
      caseSensitive: !!options.caseSensitive,
      matchCount: matches.length,
      matches: matches.slice(0, 100),
      truncated: matches.length > 100,
    },
    null,
    2
  );
}
