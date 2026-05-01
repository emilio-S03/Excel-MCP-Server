/**
 * v3.1 — read-side inspection tools.
 * All file-mode (cross-platform). For live-mode chart/pivot/shape inspection,
 * see live-inspections.ts (Windows COM + Mac AppleScript dispatcher).
 */
import { loadWorkbook, getSheet, columnNumberToLetter } from './helpers.js';

export async function getConditionalFormats(filePath: string, sheetName: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);

  const cfs: any[] = (sheet as any).conditionalFormattings ?? [];
  const rules = cfs.map((cf, i) => {
    const ranges = cf.ref ? String(cf.ref).split(/\s+/) : [];
    return {
      index: i,
      ranges,
      ruleCount: cf.rules?.length ?? 0,
      rules: (cf.rules ?? []).map((r: any) => ({
        type: r.type,
        priority: r.priority,
        operator: r.operator,
        formulae: r.formulae,
        text: r.text,
        timePeriod: r.timePeriod,
        rank: r.rank,
        percent: r.percent,
        bottom: r.bottom,
        aboveAverage: r.aboveAverage,
        equalAverage: r.equalAverage,
        stdDev: r.stdDev,
        cfvo: r.cfvo,
        color: r.color,
        style: r.style ? {
          font: r.style.font,
          fill: r.style.fill,
          border: r.style.border,
          numFmt: r.style.numFmt,
        } : undefined,
        iconSet: r.iconSet,
        showValue: r.showValue,
        reverse: r.reverse,
      })),
    };
  });

  return JSON.stringify({
    sheetName,
    totalConditionalFormats: cfs.length,
    formats: rules,
  }, null, 2);
}

export async function listDataValidations(filePath: string, sheetName: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const dvModel: Record<string, any> = (sheet as any).dataValidations?.model ?? {};

  const validations = Object.entries(dvModel).map(([range, def]: [string, any]) => ({
    range,
    type: def.type,
    operator: def.operator,
    formulae: def.formulae,
    allowBlank: def.allowBlank,
    showInputMessage: def.showInputMessage,
    promptTitle: def.promptTitle,
    prompt: def.prompt,
    showErrorMessage: def.showErrorMessage,
    errorStyle: def.errorStyle,
    errorTitle: def.errorTitle,
    error: def.error,
  }));

  return JSON.stringify({
    sheetName,
    totalValidations: validations.length,
    validations,
  }, null, 2);
}

export async function getSheetProtection(filePath: string, sheetName: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const sheetProt: any = (sheet as any).sheetProtection ?? (sheet as any).protection ?? null;
  const isProtected = !!sheetProt && (sheetProt.sheet === true || sheetProt.password !== undefined);

  return JSON.stringify({
    sheetName,
    isProtected,
    protection: sheetProt,
  }, null, 2);
}

export async function getDisplayOptions(filePath: string, sheetName: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const view: any = sheet.views?.[0] ?? {};

  const freezeCell =
    view.state === 'frozen' && (view.xSplit > 0 || view.ySplit > 0)
      ? `${columnNumberToLetter((view.xSplit ?? 0) + 1)}${(view.ySplit ?? 0) + 1}`
      : null;

  return JSON.stringify({
    sheetName,
    showGridlines: view.showGridLines !== false,
    showRowColumnHeaders: view.showRowColHeaders !== false,
    zoomLevel: view.zoomScale ?? 100,
    state: view.state ?? 'normal',
    freezePaneCell: freezeCell,
    rightToLeft: view.rightToLeft ?? false,
    activeCell: view.activeCell ?? null,
    tabColor: (sheet as any).properties?.tabColor ?? null,
  }, null, 2);
}

export async function getWorkbookProperties(filePath: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  return JSON.stringify({
    creator: workbook.creator,
    lastModifiedBy: workbook.lastModifiedBy,
    created: workbook.created,
    modified: workbook.modified,
    title: (workbook as any).title,
    subject: (workbook as any).subject,
    keywords: (workbook as any).keywords,
    category: (workbook as any).category,
    description: (workbook as any).description,
    company: (workbook as any).company,
    manager: (workbook as any).manager,
  }, null, 2);
}

export async function setWorkbookProperties(
  filePath: string,
  props: {
    creator?: string;
    lastModifiedBy?: string;
    title?: string;
    subject?: string;
    keywords?: string;
    category?: string;
    description?: string;
    company?: string;
    manager?: string;
  },
  createBackup: boolean = false
): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  if (props.creator !== undefined) workbook.creator = props.creator;
  if (props.lastModifiedBy !== undefined) workbook.lastModifiedBy = props.lastModifiedBy;
  if (props.title !== undefined) (workbook as any).title = props.title;
  if (props.subject !== undefined) (workbook as any).subject = props.subject;
  if (props.keywords !== undefined) (workbook as any).keywords = props.keywords;
  if (props.category !== undefined) (workbook as any).category = props.category;
  if (props.description !== undefined) (workbook as any).description = props.description;
  if (props.company !== undefined) (workbook as any).company = props.company;
  if (props.manager !== undefined) (workbook as any).manager = props.manager;

  workbook.modified = new Date();
  const { saveWorkbook } = await import('./helpers.js');
  await saveWorkbook(workbook, filePath, createBackup);

  return JSON.stringify({ success: true, filePath, propsSet: Object.keys(props) }, null, 2);
}

export async function getHyperlinks(filePath: string, sheetName: string): Promise<string> {
  const workbook = await loadWorkbook(filePath);
  const sheet = getSheet(workbook, sheetName);
  const links: any[] = [];

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const v = cell.value as any;
      if (v && typeof v === 'object' && (v.hyperlink || v.text)) {
        if (v.hyperlink) {
          links.push({
            address: `${columnNumberToLetter(colNumber)}${rowNumber}`,
            text: v.text ?? '',
            target: v.hyperlink,
            tooltip: v.tooltip,
          });
        }
      }
    });
  });

  return JSON.stringify({ sheetName, totalHyperlinks: links.length, hyperlinks: links }, null, 2);
}
