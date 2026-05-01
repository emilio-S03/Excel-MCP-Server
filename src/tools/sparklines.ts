/**
 * v3.1 — sparkline management.
 *
 * Why this bypasses ExcelJS:
 *   ExcelJS has no native API for sparklines, and (worse) its read/write cycle
 *   silently strips the worksheet `<extLst>` block where sparklines live. So
 *   we do a save-via-ExcelJS to ensure we have a clean .xlsx, then re-open
 *   the .xlsx as a ZIP with JSZip and surgically inject (or remove) the
 *   sparkline group inside `xl/worksheets/sheetN.xml`.
 *
 * Sparkline OOXML reference (ECMA-376 Part 1, §18.3.1.27.1, x14 namespace):
 *   <extLst>
 *     <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B6A02}"
 *          xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
 *       <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
 *         <x14:sparklineGroup type="line" displayEmptyCellsAs="gap" markers="1" ...>
 *           <x14:colorSeries rgb="FF376092"/>
 *           <x14:colorNegative rgb="FFFF0000"/>
 *           <x14:sparklines>
 *             <x14:sparkline>
 *               <xm:f>Sheet1!A1:A10</xm:f>
 *               <xm:sqref>B1</xm:sqref>
 *             </x14:sparkline>
 *           </x14:sparklines>
 *         </x14:sparklineGroup>
 *       </x14:sparklineGroups>
 *     </ext>
 *   </extLst>
 */
import JSZip from 'jszip';
import { promises as fs } from 'fs';
import {
  loadWorkbook,
  getSheet,
  saveWorkbook,
  ensureFilePathAllowed,
  parseRange,
  columnLetterToNumber,
  columnNumberToLetter,
} from './helpers.js';

export interface AddSparklineOptions {
  type: 'line' | 'column' | 'winLoss';
  dataRange: string;          // e.g., "A1:A10" or "Sheet1!A1:A10"
  locationRange: string;      // e.g., "B1" (single cell) or "B1:B5" (one sparkline per cell)
  color?: string;             // hex color for the series, e.g., "#376092" or "376092"
  negativeColor?: string;     // hex color for negative values
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  createBackup?: boolean;
}

const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B6A02}';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const X14AC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

function normalizeHex(c: string | undefined): string | undefined {
  if (!c) return undefined;
  let s = c.trim();
  if (s.startsWith('#')) s = s.slice(1);
  if (s.length === 6) s = 'FF' + s;
  if (!/^[0-9A-Fa-f]{8}$/.test(s)) {
    throw new Error(`Invalid color hex (expected #RRGGBB or #AARRGGBB): ${c}`);
  }
  return s.toUpperCase();
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Resolve a sheet name to its ZIP path (xl/worksheets/sheetN.xml) by
 * reading workbook.xml + workbook.xml.rels. ExcelJS does not always number
 * sheets in display order (e.g., after reordering / hiding), so we cannot
 * just use the sheet's display index.
 */
async function resolveSheetPart(zip: JSZip, sheetName: string): Promise<{ path: string; xml: string }> {
  const wbFile = zip.file('xl/workbook.xml');
  if (!wbFile) throw new Error('Corrupt .xlsx: missing xl/workbook.xml');
  const wbXml = await wbFile.async('string');

  // Match each <sheet ... /> (attribute values may contain '/' so match
  // up to the closing '/>' rather than excluding '/' from the attr region).
  const sheetRe = /<sheet\b[^>]*?\/>/g;
  let sheetRid: string | null = null;
  for (const m of wbXml.match(sheetRe) ?? []) {
    const nameMatch = m.match(/\bname="([^"]+)"/);
    const ridMatch = m.match(/\b(?:r:id|r:Id|relationships:id)="([^"]+)"/);
    if (nameMatch && ridMatch && nameMatch[1] === sheetName) {
      sheetRid = ridMatch[1];
      break;
    }
  }
  if (!sheetRid) throw new Error(`Sheet not found in workbook.xml: ${sheetName}`);

  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!relsFile) throw new Error('Corrupt .xlsx: missing xl/_rels/workbook.xml.rels');
  const relsXml = await relsFile.async('string');

  const relRe = /<Relationship\b[^>]*?\/>/g;
  let target: string | null = null;
  for (const m of relsXml.match(relRe) ?? []) {
    const idMatch = m.match(/\bId="([^"]+)"/);
    const tgtMatch = m.match(/\bTarget="([^"]+)"/);
    if (idMatch && tgtMatch && idMatch[1] === sheetRid) {
      target = tgtMatch[1];
      break;
    }
  }
  if (!target) throw new Error(`Could not resolve sheet relationship: ${sheetRid}`);

  // Targets are typically "worksheets/sheet1.xml" relative to xl/
  const path = target.startsWith('/')
    ? target.slice(1)
    : `xl/${target.replace(/^\.\//, '')}`;

  const sheetFile = zip.file(path);
  if (!sheetFile) throw new Error(`Sheet part not found in zip: ${path}`);
  const xml = await sheetFile.async('string');
  return { path, xml };
}

/**
 * Build the qualified data range reference string used in <xm:f>.
 * If the dataRange already includes a sheet name (Sheet1!A1:A10) keep it,
 * otherwise prefix with the current sheet name (quoted if needed).
 */
function qualifyRangeRef(rangeRef: string, sheetName: string): string {
  if (rangeRef.includes('!')) return rangeRef;
  const needsQuote = /[\s'!]/.test(sheetName);
  const quoted = needsQuote ? `'${sheetName.replace(/'/g, "''")}'` : sheetName;
  return `${quoted}!${rangeRef}`;
}

/**
 * Expand a location range into individual cell addresses, one per sparkline.
 * Single-cell location ("B1") -> ["B1"]. A column range ("B1:B5") -> 5 cells.
 * If the dataRange covers multiple rows or columns and locationRange is a
 * single cell, we still produce one sparkline (the whole range goes into one
 * cell).
 */
function expandLocationCells(locationRange: string): string[] {
  if (/^[A-Z]+\d+$/.test(locationRange)) {
    return [locationRange];
  }
  const parsed = parseRange(locationRange);
  const cells: string[] = [];
  for (let r = parsed.startRow; r <= parsed.endRow; r++) {
    for (let c = parsed.startCol; c <= parsed.endCol; c++) {
      cells.push(`${columnNumberToLetter(c)}${r}`);
    }
  }
  return cells;
}

/**
 * If the user provides a location range with multiple cells AND a multi-row
 * or multi-column data range, slice the data range so each location cell gets
 * its own row/column from the data range. Mirrors Excel's "Insert Sparkline"
 * behavior when you select a data range and a same-shaped location range.
 */
function deriveSparklineRefs(
  dataRange: string,
  locationCells: string[],
  contextSheetName: string,
): { dataRef: string; cellRef: string }[] {
  // Strip any sheet prefix to inspect the bare A1:B2 form.
  let bareRange = dataRange;
  let sheetPrefix: string | null = null;
  const bangIdx = dataRange.lastIndexOf('!');
  if (bangIdx >= 0) {
    sheetPrefix = dataRange.slice(0, bangIdx);
    bareRange = dataRange.slice(bangIdx + 1);
  }

  // Single location cell — one sparkline covering the whole data range.
  if (locationCells.length === 1) {
    return [{ dataRef: qualifyRangeRef(dataRange, contextSheetName), cellRef: locationCells[0] }];
  }

  // For multi-cell location, try to slice the data range row-by-row or
  // column-by-column so each output cell maps to a slice.
  if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(bareRange)) {
    // Can't slice — fall back to giving every cell the full range.
    return locationCells.map((c) => ({ dataRef: qualifyRangeRef(dataRange, contextSheetName), cellRef: c }));
  }
  const parsed = parseRange(bareRange);
  const dataRows = parsed.endRow - parsed.startRow + 1;
  const dataCols = parsed.endCol - parsed.startCol + 1;

  const buildRef = (startCol: number, startRow: number, endCol: number, endRow: number) => {
    const ref = `${columnNumberToLetter(startCol)}${startRow}:${columnNumberToLetter(endCol)}${endRow}`;
    const fq = sheetPrefix ? `${sheetPrefix}!${ref}` : ref;
    return qualifyRangeRef(fq, contextSheetName);
  };

  // If location is a vertical strip and the data range has matching rows, slice by row.
  if (dataRows === locationCells.length) {
    return locationCells.map((cell, idx) => ({
      dataRef: buildRef(parsed.startCol, parsed.startRow + idx, parsed.endCol, parsed.startRow + idx),
      cellRef: cell,
    }));
  }
  // If location is horizontal and data has matching columns, slice by column.
  if (dataCols === locationCells.length) {
    return locationCells.map((cell, idx) => ({
      dataRef: buildRef(parsed.startCol + idx, parsed.startRow, parsed.startCol + idx, parsed.endRow),
      cellRef: cell,
    }));
  }
  // Shape mismatch — every output cell gets the full data range.
  return locationCells.map((c) => ({ dataRef: qualifyRangeRef(dataRange, contextSheetName), cellRef: c }));
}

function buildSparklineGroupXml(
  options: AddSparklineOptions,
  sheetName: string,
): string {
  const refs = deriveSparklineRefs(options.dataRange, expandLocationCells(options.locationRange), sheetName);
  const color = normalizeHex(options.color) ?? 'FF376092';
  const negColor = normalizeHex(options.negativeColor) ?? 'FFFF0000';

  // Group attributes
  const attrs: string[] = [];
  if (options.type === 'column') attrs.push('type="column"');
  else if (options.type === 'winLoss') attrs.push('type="stacked"'); // OOXML name for win/loss
  // line is the default — no attribute needed.
  attrs.push('displayEmptyCellsAs="gap"');
  if (options.markers) attrs.push('markers="1"');
  if (options.high) attrs.push('high="1"');
  if (options.low) attrs.push('low="1"');
  if (options.first) attrs.push('first="1"');
  if (options.last) attrs.push('last="1"');

  const groupOpen = `<x14:sparklineGroup ${attrs.join(' ')}>`;

  const colorBlocks: string[] = [];
  colorBlocks.push(`<x14:colorSeries rgb="${color}"/>`);
  colorBlocks.push(`<x14:colorNegative rgb="${negColor}"/>`);
  colorBlocks.push(`<x14:colorAxis rgb="FF000000"/>`);
  colorBlocks.push(`<x14:colorMarkers rgb="FFD00000"/>`);
  colorBlocks.push(`<x14:colorFirst rgb="FFD00000"/>`);
  colorBlocks.push(`<x14:colorLast rgb="FFD00000"/>`);
  colorBlocks.push(`<x14:colorHigh rgb="FFD00000"/>`);
  colorBlocks.push(`<x14:colorLow rgb="FFD00000"/>`);

  const sparkLines = refs
    .map(
      (r) =>
        `<x14:sparkline><xm:f>${escapeXml(r.dataRef)}</xm:f><xm:sqref>${escapeXml(r.cellRef)}</xm:sqref></x14:sparkline>`,
    )
    .join('');

  return (
    groupOpen +
    colorBlocks.join('') +
    `<x14:sparklines>${sparkLines}</x14:sparklines>` +
    `</x14:sparklineGroup>`
  );
}

/**
 * Inject a sparkline group into the worksheet XML.
 *
 * If <extLst> exists with a sparkline ext block, append a new
 * <x14:sparklineGroup> inside the existing <x14:sparklineGroups>.
 * Otherwise create the full extLst → ext → sparklineGroups → sparklineGroup chain.
 *
 * Worksheet-level extLst MUST be the very last child of <worksheet> per
 * the schema. We insert it just before </worksheet>.
 */
function injectSparklineGroup(sheetXml: string, groupXml: string): string {
  // Ensure worksheet root declares the x14ac/mc namespaces so consumers parse
  // the ext correctly. (Not strictly required for the ext itself, but Excel
  // is happiest when these are present once a worksheet uses extLst.)
  let xml = sheetXml;
  const worksheetOpenMatch = xml.match(/<worksheet\b[^>]*>/);
  if (worksheetOpenMatch) {
    let openTag = worksheetOpenMatch[0];
    let changed = false;
    if (!openTag.includes(`xmlns:mc=`)) {
      openTag = openTag.replace(/<worksheet\b/, `<worksheet xmlns:mc="${MC_NS}"`);
      changed = true;
    }
    if (!openTag.includes(`xmlns:x14ac=`)) {
      openTag = openTag.replace(/<worksheet\b/, `<worksheet xmlns:x14ac="${X14AC_NS}"`);
      changed = true;
    }
    if (changed) xml = xml.replace(worksheetOpenMatch[0], openTag);
  }

  // Locate (or create) the sparkline ext block.
  const sparklineExtRe = new RegExp(
    `<ext\\b[^>]*uri="\\{05C60535-1F16-4fd2-B633-F4F36F0B6A02\\}"[^>]*>([\\s\\S]*?)</ext>`,
    'i',
  );
  const sparklineExtMatch = xml.match(sparklineExtRe);
  if (sparklineExtMatch) {
    // Append into existing <x14:sparklineGroups>...
    const groupsCloseRe = /<\/x14:sparklineGroups>/;
    if (groupsCloseRe.test(sparklineExtMatch[0])) {
      const updated = sparklineExtMatch[0].replace(
        groupsCloseRe,
        `${groupXml}</x14:sparklineGroups>`,
      );
      return xml.replace(sparklineExtMatch[0], updated);
    }
    // ext exists but no groups element (very unusual) — wrap and inject.
    const updated = sparklineExtMatch[0].replace(
      /<\/ext>$/,
      `<x14:sparklineGroups xmlns:xm="${XM_NS}">${groupXml}</x14:sparklineGroups></ext>`,
    );
    return xml.replace(sparklineExtMatch[0], updated);
  }

  // No sparkline ext — build the full chain.
  const newExtBlock =
    `<ext xmlns:x14="${X14_NS}" uri="${SPARKLINE_EXT_URI}">` +
    `<x14:sparklineGroups xmlns:xm="${XM_NS}">${groupXml}</x14:sparklineGroups>` +
    `</ext>`;

  // If <extLst> already exists, insert into it.
  if (/<extLst>/.test(xml)) {
    return xml.replace(/<\/extLst>/, `${newExtBlock}</extLst>`);
  }
  // No extLst — must be inserted as the last child of <worksheet>.
  return xml.replace(/<\/worksheet>\s*$/, `<extLst>${newExtBlock}</extLst></worksheet>`);
}

function cellInRange(cellAddr: string, rangeOrCell: string): boolean {
  if (cellAddr === rangeOrCell) return true;
  if (!/:/.test(rangeOrCell)) return cellAddr === rangeOrCell;
  const parsed = parseRange(rangeOrCell);
  const cellMatch = cellAddr.match(/^([A-Z]+)(\d+)$/);
  if (!cellMatch) return false;
  const col = columnLetterToNumber(cellMatch[1]);
  const row = parseInt(cellMatch[2], 10);
  return (
    col >= parsed.startCol &&
    col <= parsed.endCol &&
    row >= parsed.startRow &&
    row <= parsed.endRow
  );
}

/**
 * Remove any sparkline groups whose <x14:sparkline> entries cover one of the
 * locationCells. If locationCells is empty, remove ALL sparkline groups on
 * the sheet.
 */
function removeSparklineGroups(
  sheetXml: string,
  locationCells: string[],
): { xml: string; removedGroups: number; removedSparklines: number } {
  const sparklineExtRe = new RegExp(
    `<ext\\b[^>]*uri="\\{05C60535-1F16-4fd2-B633-F4F36F0B6A02\\}"[^>]*>([\\s\\S]*?)</ext>`,
    'i',
  );
  const sparklineExtMatch = sheetXml.match(sparklineExtRe);
  if (!sparklineExtMatch) {
    return { xml: sheetXml, removedGroups: 0, removedSparklines: 0 };
  }

  const extInner = sparklineExtMatch[1];
  const groupsRe = /<x14:sparklineGroup\b[\s\S]*?<\/x14:sparklineGroup>/g;
  const groups = extInner.match(groupsRe) ?? [];

  let removedGroups = 0;
  let removedSparklines = 0;
  const keptGroups: string[] = [];

  for (const group of groups) {
    const sparkRe = /<x14:sparkline\b[\s\S]*?<\/x14:sparkline>/g;
    const sparks = group.match(sparkRe) ?? [];
    let groupTouched = false;
    let keptSparks: string[] = [];
    for (const spark of sparks) {
      const sqrefMatch = spark.match(/<xm:sqref>([^<]+)<\/xm:sqref>/);
      const sqref = sqrefMatch ? sqrefMatch[1].trim() : '';
      const matchesAny =
        locationCells.length === 0 ||
        locationCells.some((cell) => cellInRange(cell, sqref));
      if (matchesAny) {
        groupTouched = true;
        removedSparklines++;
      } else {
        keptSparks.push(spark);
      }
    }
    if (!groupTouched) {
      keptGroups.push(group);
      continue;
    }
    if (keptSparks.length === 0) {
      removedGroups++;
      continue;
    }
    // Some sparklines kept — rebuild the group with only those.
    const rebuilt = group.replace(
      /<x14:sparklines>[\s\S]*<\/x14:sparklines>/,
      `<x14:sparklines>${keptSparks.join('')}</x14:sparklines>`,
    );
    keptGroups.push(rebuilt);
  }

  let newXml: string;
  if (keptGroups.length === 0) {
    // Drop the entire sparkline ext block.
    let withoutExt = sheetXml.replace(sparklineExtMatch[0], '');
    // If extLst is now empty, drop it too.
    withoutExt = withoutExt.replace(/<extLst>\s*<\/extLst>/, '');
    newXml = withoutExt;
  } else {
    const newInner = extInner.replace(
      /<x14:sparklineGroups\b[^>]*>[\s\S]*<\/x14:sparklineGroups>/,
      (m) => {
        const openTag = m.match(/<x14:sparklineGroups\b[^>]*>/)?.[0] ?? '<x14:sparklineGroups>';
        return `${openTag}${keptGroups.join('')}</x14:sparklineGroups>`;
      },
    );
    const newExt = sparklineExtMatch[0].replace(extInner, newInner);
    newXml = sheetXml.replace(sparklineExtMatch[0], newExt);
  }

  return { xml: newXml, removedGroups, removedSparklines };
}

/**
 * Round-trip a workbook through ExcelJS save (to normalize any pending edits
 * — though for sparklines we typically just call this on an existing file —
 * then operate on the .xlsx as a ZIP).
 *
 * If `roundTripFirst` is false, we open the file directly as a zip without
 * an ExcelJS save first. This preserves any extLst already on disk (the
 * common case for `removeSparklines` and for `addSparkline` when called on a
 * file the caller has not just modified).
 */
async function withSheetZip(
  filePath: string,
  sheetName: string,
  mutate: (xml: string) => { newXml: string; meta?: Record<string, unknown> },
  options: { createBackup?: boolean; roundTripFirst?: boolean } = {},
): Promise<{ meta?: Record<string, unknown> }> {
  ensureFilePathAllowed(filePath);

  if (options.roundTripFirst) {
    // Validate workbook + sheet exist via ExcelJS, then rewrite to disk.
    const wb = await loadWorkbook(filePath);
    getSheet(wb, sheetName); // throws if missing
    await saveWorkbook(wb, filePath, options.createBackup ?? false);
  } else if (options.createBackup) {
    try {
      await fs.access(filePath);
      await fs.copyFile(filePath, `${filePath}.backup`);
    } catch {
      /* nothing to back up */
    }
  }

  const buf = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(buf);
  const { path, xml } = await resolveSheetPart(zip, sheetName);
  const { newXml, meta } = mutate(xml);
  zip.file(path, newXml);

  // Mirror ExcelJS's compression so the output stays a valid .xlsx.
  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await fs.writeFile(filePath, out);

  return { meta };
}

export async function addSparkline(
  filePath: string,
  sheetName: string,
  options: AddSparklineOptions,
): Promise<string> {
  if (!options || !options.type || !options.dataRange || !options.locationRange) {
    throw new Error('addSparkline requires type, dataRange, and locationRange');
  }
  const groupXml = buildSparklineGroupXml(options, sheetName);

  const { meta } = await withSheetZip(
    filePath,
    sheetName,
    (xml) => ({ newXml: injectSparklineGroup(xml, groupXml) }),
    { createBackup: options.createBackup, roundTripFirst: false },
  );

  return JSON.stringify(
    {
      success: true,
      filePath,
      sheetName,
      type: options.type,
      dataRange: options.dataRange,
      locationRange: options.locationRange,
      sparklineCount: expandLocationCells(options.locationRange).length,
      ...meta,
    },
    null,
    2,
  );
}

export async function removeSparklines(
  filePath: string,
  sheetName: string,
  locationRange?: string,
  options: { createBackup?: boolean } = {},
): Promise<string> {
  ensureFilePathAllowed(filePath);

  const locationCells = locationRange ? expandLocationCells(locationRange) : [];

  let removedGroups = 0;
  let removedSparklines = 0;

  await withSheetZip(
    filePath,
    sheetName,
    (xml) => {
      const result = removeSparklineGroups(xml, locationCells);
      removedGroups = result.removedGroups;
      removedSparklines = result.removedSparklines;
      return { newXml: result.xml };
    },
    { createBackup: options.createBackup, roundTripFirst: false },
  );

  return JSON.stringify(
    {
      success: true,
      filePath,
      sheetName,
      locationRange: locationRange ?? null,
      removedGroups,
      removedSparklines,
    },
    null,
    2,
  );
}
