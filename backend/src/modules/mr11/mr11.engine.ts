import { PrismaClient, RoleCode } from '@prisma/client';
import { MR11_ORDERED_COLUMNS, MR11_SOURCE_KEY_MAP, ORDERED_HEADER_LIST } from '../../config/mr11.config';
import { FortuneSheet } from '../../utils/excel-normalizer';

export const Mr11Status = {
  EMPTY: 'EMPTY',
  PARTIAL: 'PARTIAL',
  READY: 'READY',
  FAILED: 'FAILED',
} as const;

export type Mr11Status = (typeof Mr11Status)[keyof typeof Mr11Status];

interface ExtractedRow {
  data: Record<string, any>;
  fontColor?: string;
  fillColor?: string;
  rawCells?: Record<number, any>;
}

function parseNumeric(val: any): number {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function cleanStr(val: any): string {
  if (val === null || val === undefined) return '';
  return String(val)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normalizeStream(val: any): string {
  if (val === null || val === undefined || val === '') return '1';
  const num = parseFloat(String(val).replace(/[^0-9.]/g, ''));
  return isNaN(num) ? String(val).trim().toUpperCase() : String(num);
}

function normalizeColor(color: any): string {
  if (!color) return '#000000';
  const s = String(color).trim().toUpperCase();
  if (s === 'BLACK' || s === '#000' || s === '#000000') return '#000000';
  if (s === 'WHITE' || s === '#FFF' || s === '#FFFFFF') return '#FFFFFF';
  return s.startsWith('#') ? s : `#${s}`;
}

function normalizeFillColor(color: any): string {
  if (!color) return '';
  const s = String(color).trim().toUpperCase();
  if (
    s === '#FFFFFF' ||
    s === 'WHITE' ||
    s === '#000000' ||
    s === 'BLACK' ||
    s === 'TRANSPARENT' ||
    s === 'NONE'
  ) {
    return '';
  }
  return s.startsWith('#') ? s : `#${s}`;
}

function getProjectIdentifier(row: Record<string, any>): string {
  const pNo = cleanStr(
    row['Project No'] ||
    row['Project No.'] ||
    row['Project No. (from design column A)'] ||
    row['Project No. (from bd column B)'] ||
    row['PROJECT NO']
  );
  const pShort = cleanStr(
    row['Short Name'] ||
    row['Project Shortname'] ||
    row['Project Shortname (from bd column C)'] ||
    row['Project Shortname (from planning column B)'] ||
    row['Shortname']
  );
  const pName = cleanStr(
    row['Customer & Project Name'] ||
    row['Project Name'] ||
    row['Project Name (from design column B)'] ||
    row['Project Name (from bd column A)'] ||
    row['Project Name (from planning column C)']
  );
  return pNo || pShort || pName;
}

function parseFlexibleDate(val: any): Date | null {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  const rawStr = String(val).trim();
  if (!rawStr || rawStr === '-' || rawStr === '0' || rawStr.toLowerCase() === 'null') return null;

  const num = typeof val === 'number' ? val : parseFloat(rawStr);
  if (!isNaN(num) && num > 30000 && !rawStr.includes('/') && !rawStr.includes('-')) {
    const parsedExcelDate = new Date(Math.round((num - 25569) * 86400 * 1000));
    return isNaN(parsedExcelDate.getTime()) ? null : parsedExcelDate;
  }

  const dmyMatch = rawStr.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    let year = parseInt(dmyMatch[3], 10);
    if (year < 100) year += 2000;
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
  }

  const ymdMatch = rawStr.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10) - 1;
    const day = parseInt(ymdMatch[3], 10);
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
  }

  const parsed = new Date(rawStr);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateString(val: any): string | null {
  if (!val) return null;
  const d = parseFlexibleDate(val);
  if (d) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return String(val).trim();
}

function resolveConsolidatedDesignStatus(statuses: string[]): string {
  if (statuses.length === 0) return 'Not Start';
  const normalized = statuses.map((s) => s.toLowerCase().trim());
  const hasCompleted = normalized.some((s) => s.includes('complete'));
  const hasOngoing = normalized.some((s) => s.includes('ongoing') || s.includes('progress') || s.includes('in progress'));
  const hasNotStarted = normalized.some((s) => s.includes('start') || s.includes('pending') || s.includes('hold'));

  if (hasOngoing || (hasCompleted && hasNotStarted)) return 'Ongoing';
  if (hasCompleted && !hasNotStarted && !hasOngoing) return 'Completed';
  return 'Not Start';
}

function resolveConsolidatedShellplanStatus(statuses: string[]): string {
  if (statuses.length === 0) return 'Not Started';
  const normalized = statuses.map((s) => s.toLowerCase().trim());

  if (normalized.some((s) => s.includes('progress') || s.includes('ongoing'))) {
    return 'In Progress';
  }
  if (normalized.some((s) => s.includes('pending'))) {
    return 'Pending Consultant Drawings';
  }
  const allApproved = normalized.every((s) => s.includes('approve'));
  if (allApproved) {
    return 'Approved';
  }
  const allNotStarted = normalized.every((s) => s.includes('start') || s.includes('haven') || s.includes('hold'));
  if (allNotStarted) {
    return 'Not Started';
  }
  return normalized[0] ? normalized[0].charAt(0).toUpperCase() + normalized[0].slice(1) : 'Not Started';
}

function isCellFilled(val: any): boolean {
  if (val === null || val === undefined) return false;
  const s = String(val).trim();
  if (s === '' || s === '-' || s === '0' || s === '0.0' || s === '0.00') return false;
  const num = parseNumeric(val);
  return !isNaN(num) ? num > 0 : s.length > 0;
}

export function sheetToRecordsWithStyles(sheet: FortuneSheet): {
  rows: ExtractedRow[];
  detectedSeries: number;
  headers: Record<number, string>;
} {
  const headers: Record<number, string> = {};
  const rowsMap: Record<number, Record<string, any>> = {};
  const rawCellsMap: Record<number, Record<number, any>> = {};
  const rowStyleMap: Record<number, { fontColor?: string; fillColor?: string }> = {};
  let detectedSeries = 0;

  if (!sheet) return { rows: [], detectedSeries: 0, headers: {} };

  const celldata = Array.isArray(sheet.celldata) ? sheet.celldata : [];

  if (celldata.length === 0 && Array.isArray((sheet as any).data)) {
    const matrix: any[][] = (sheet as any).data;
    if (matrix.length > 0 && Array.isArray(matrix[0])) {
      for (let hR = 0; hR <= Math.min(3, matrix.length - 1); hR++) {
        if (Array.isArray(matrix[hR])) {
          matrix[hR].forEach((col: any, cIdx: number) => {
            const val = typeof col === 'object' && col !== null ? col?.v ?? col?.m : col;
            if (val !== undefined && val !== null && !headers[cIdx]) {
              headers[cIdx] = String(val).trim();
            }
          });
        }
      }

      for (let r = 1; r < matrix.length; r++) {
        const row = matrix[r];
        if (!row || !Array.isArray(row)) continue;
        rowsMap[r] = {};
        rawCellsMap[r] = {};
        rowStyleMap[r] = {};
        row.forEach((cell: any, cIdx: number) => {
          const colName = headers[cIdx];
          const val = typeof cell === 'object' && cell !== null ? cell?.v ?? cell?.m : cell;
          if (val !== undefined) rawCellsMap[r][cIdx] = val;
          if (colName && val !== undefined) rowsMap[r][colName] = val;

          if (cIdx === 10 && val !== undefined) rowsMap[r]['__COLUMN_K__'] = val;
          if (cIdx === 16 && val !== undefined) rowsMap[r]['__COLUMN_Q__'] = val;
          if (cIdx === 22 && val !== undefined) rowsMap[r]['__COLUMN_W__'] = val;

          if (cell && typeof cell === 'object' && cIdx <= 4) {
            const fc = cell.fc || cell.v?.fc;
            const bg = cell.bg || cell.v?.bg;
            if (fc && !rowStyleMap[r].fontColor) {
              const sfc = String(fc).toUpperCase();
              if (sfc !== '#000000' && sfc !== 'BLACK' && sfc !== '#000') {
                rowStyleMap[r].fontColor = sfc;
              }
            }
            if (bg && !rowStyleMap[r].fillColor) {
              const sbg = String(bg).toUpperCase();
              if (sbg !== '#FFFFFF' && sbg !== 'WHITE' && sbg !== '#000000' && sbg !== 'TRANSPARENT') {
                rowStyleMap[r].fillColor = sbg;
              }
            }
          }
        });
      }
    }
  } else {
    for (let checkR = 0; checkR <= 3; checkR++) {
      const rCells = celldata.filter((c) => c && c.r === checkR);
      for (const cell of rCells) {
        const text = String(cell.v?.v ?? cell.v?.m ?? '').trim();
        if (text) {
          if (!headers[cell.c]) headers[cell.c] = text;
          const match = text.match(/series\s*(\d+)/i);
          if (match && detectedSeries === 0) {
            detectedSeries = parseInt(match[1], 10);
          }
        }
      }
    }

    const bodyCells = celldata.filter((c) => c && c.r > 0);
    for (const cell of bodyCells) {
      if (!rowsMap[cell.r]) {
        rowsMap[cell.r] = {};
        rawCellsMap[cell.r] = {};
        rowStyleMap[cell.r] = {};
      }

      const val = cell.v?.m !== undefined && cell.v?.m !== null && String(cell.v.m).trim() !== ''
        ? cell.v.m
        : cell.v?.v ?? cell.v?.m;

      if (val !== undefined) rawCellsMap[cell.r][cell.c] = val;

      const colName = headers[cell.c];
      if (colName && val !== undefined) {
        rowsMap[cell.r][colName] = val;
      }

      if (cell.c === 10 && val !== undefined) rowsMap[cell.r]['__COLUMN_K__'] = val;
      if (cell.c === 16 && val !== undefined) rowsMap[cell.r]['__COLUMN_Q__'] = val;
      if (cell.c === 22 && val !== undefined) rowsMap[cell.r]['__COLUMN_W__'] = val;

      if (cell.c <= 4) {
        if (cell.v?.fc && !rowStyleMap[cell.r].fontColor) {
          const fc = String(cell.v.fc).toUpperCase();
          if (fc !== '#000000' && fc !== 'BLACK' && fc !== '#000') {
            rowStyleMap[cell.r].fontColor = fc;
          }
        }
        if (cell.v?.bg && !rowStyleMap[cell.r].fillColor) {
          const bg = String(cell.v.bg).toUpperCase();
          if (bg !== '#FFFFFF' && bg !== 'WHITE' && bg !== '#000000' && bg !== 'TRANSPARENT') {
            rowStyleMap[cell.r].fillColor = bg;
          }
        }
      }
    }
  }

  const sortedRowKeys = Object.keys(rowsMap).map(Number).sort((a, b) => a - b);

  let lastProjectNo: any = null;
  let lastShortname: any = null;
  let lastName: any = null;
  let lastFontColor: string | undefined = undefined;
  let lastFillColor: string | undefined = undefined;

  for (const rKey of sortedRowKeys) {
    const row = rowsMap[rKey];
    const pNo = row['Project No'] || row['Project No.'] || row['Project No. (from design column A)'] || row['Project No. (from bd column B)'];
    const pShort = row['Short Name'] || row['Project Shortname'] || row['Project Shortname (from bd column C)'];
    const pName = row['Customer & Project Name'] || row['Project Name'] || row['Project Name (from design column B)'] || row['Project Name (from bd column A)'];

    if (pNo || pShort || pName) {
      lastProjectNo = pNo || null;
      lastShortname = pShort || null;
      lastName = pName || null;
      lastFontColor = rowStyleMap[rKey]?.fontColor;
      lastFillColor = rowStyleMap[rKey]?.fillColor;
    } else {
      if (lastProjectNo && !row['Project No']) row['Project No'] = lastProjectNo;
      if (lastShortname && !row['Short Name']) row['Short Name'] = lastShortname;
      if (lastName && !row['Customer & Project Name']) row['Customer & Project Name'] = lastName;
      if (lastFontColor && !rowStyleMap[rKey]?.fontColor) {
        rowStyleMap[rKey].fontColor = lastFontColor;
      }
      if (lastFillColor && !rowStyleMap[rKey]?.fillColor) {
        rowStyleMap[rKey].fillColor = lastFillColor;
      }
    }
  }

  const rows = sortedRowKeys.map((rKey) => ({
    data: rowsMap[rKey],
    fontColor: normalizeColor(rowStyleMap[rKey]?.fontColor),
    fillColor: normalizeFillColor(rowStyleMap[rKey]?.fillColor),
    rawCells: rawCellsMap[rKey],
  }));

  return { rows, detectedSeries, headers };
}

function findCellValue(row: Record<string, any>, candidateHeader: string): any {
  if (!row) return null;
  if (row[candidateHeader] !== undefined) return row[candidateHeader];
  const target = candidateHeader.toLowerCase().trim();
  for (const key of Object.keys(row)) {
    const normalized = key.toLowerCase().trim();
    if (normalized === target || normalized.startsWith(target) || target.startsWith(normalized)) {
      return row[key];
    }
  }
  return null;
}

export async function executeMr11Pipeline(prisma: PrismaClient): Promise<string> {
  const activeDepartments = await prisma.department.findMany({
    include: { activeVersion: true },
  });

  const sourceSnapshot: Record<string, any> = {};
  const datasetMap: Partial<Record<RoleCode, ExtractedRow[]>> = {};
  let productionHeaders: Record<number, string> = {};
  let dispatchHeaders: Record<number, string> = {};
  let detectedProdSeries = 0;

  for (const dept of activeDepartments) {
    if (dept.activeVersion?.parsedWorkbook) {
      sourceSnapshot[dept.code] = dept.activeVersion.id;
      const rawWb = dept.activeVersion.parsedWorkbook;

      const sheets: FortuneSheet[] = Array.isArray(rawWb)
        ? (rawWb as unknown as FortuneSheet[])
        : (((rawWb as any)?.sheets as unknown as FortuneSheet[]) || []);

      if (sheets.length > 0) {
        const { rows, detectedSeries, headers } = sheetToRecordsWithStyles(sheets[0]);
        datasetMap[dept.code] = rows;
        if (dept.code === RoleCode.PRODUCTION) {
          productionHeaders = headers;
          if (detectedSeries > 0) {
            detectedProdSeries = detectedSeries;
          }
        }
        if (dept.code === RoleCode.DISPATCH) {
          dispatchHeaders = headers;
        }
      }
    }
  }

  const STREAM_HEADER_CANDIDATES = [
    'Stream',
    'stream',
    'Stream (from design column D)',
    'Stream (from planning column D)',
    'stream (from bd column D)',
    'Stream (from bd column D)',
  ];

  const todayStr = new Date().toISOString().split('T')[0];

  // --------------------------------------------------------------------------
  // 1. PLANNING SERIES & QUANTITY TRACKER
  // --------------------------------------------------------------------------
  const planningRows = datasetMap[RoleCode.PLANNING] || [];
  const incomingPlanningTotals: Record<string, { projectNo: string; pShort: string; pName: string; stream: string; fontColor: string; totalQty: number }> = {};

  for (const row of planningRows) {
    const pNo = cleanStr(findCellValue(row.data, 'Project No') || findCellValue(row.data, 'Project No.') || findCellValue(row.data, 'Project No. (from design column A)'));
    const pShort = cleanStr(findCellValue(row.data, 'Short Name') || findCellValue(row.data, 'Project Shortname') || findCellValue(row.data, 'Project Shortname (from bd column C)'));
    const pName = findCellValue(row.data, 'Customer & Project Name') || findCellValue(row.data, 'Project Name') || findCellValue(row.data, 'Project Name (from design column B)');

    let stream = '1';
    for (const sh of STREAM_HEADER_CANDIDATES) {
      const v = findCellValue(row.data, sh);
      if (v !== null && v !== undefined && v !== '') {
        stream = normalizeStream(v);
        break;
      }
    }

    const rawSeries = findCellValue(row.data, 'Series') || findCellValue(row.data, 'Series No');
    const seriesNumber = rawSeries ? parseInt(String(rawSeries).replace(/[^0-9]/g, ''), 10) : 0;
    const fontColor = normalizeColor(row.fontColor);

    if (pNo && stream) {
      const compKey = `${pNo}_${stream}_${fontColor}`;
      const rawQty = findCellValue(row.data, 'Total Quantity (m2)');
      const parsedQty = rawQty ? parseFloat(String(rawQty).replace(/[^0-9.-]/g, '')) || 0 : 0;

      if (!incomingPlanningTotals[compKey]) {
        incomingPlanningTotals[compKey] = {
          projectNo: pNo,
          pShort,
          pName,
          stream,
          fontColor,
          totalQty: parsedQty,
        };
      } else {
        incomingPlanningTotals[compKey].totalQty += parsedQty;
      }

      if (seriesNumber > 0) {
        const rawProcessed = findCellValue(row.data, 'Total Processed ') || findCellValue(row.data, 'Total Processed');
        const totalProcessed = parseNumeric(rawProcessed);
        const closingDate = findCellValue(row.data, 'Closing Date ')
          ? String(findCellValue(row.data, 'Closing Date '))
          : findCellValue(row.data, 'Closing Date')
          ? String(findCellValue(row.data, 'Closing Date'))
          : findCellValue(row.data, 'Processed Date')
          ? String(findCellValue(row.data, 'Processed Date'))
          : null;

        await prisma.planningSeriesHistory.upsert({
          where: {
            projectNo_stream_fontColor_seriesNumber: {
              projectNo: pNo,
              stream,
              fontColor,
              seriesNumber,
            },
          },
          update: {
            totalProcessed,
            totalQuantity: parsedQty > 0 ? parsedQty : null,
            closingDate,
            projectShortname: pShort || null,
            projectName: pName || null,
          },
          create: {
            projectNo: pNo,
            projectShortname: pShort || null,
            projectName: pName || null,
            stream,
            fontColor,
            seriesNumber,
            totalProcessed,
            totalQuantity: parsedQty > 0 ? parsedQty : null,
            closingDate,
          },
        });
      }
    }
  }

  for (const key of Object.keys(incomingPlanningTotals)) {
    const item = incomingPlanningTotals[key];
    const existing = await prisma.planningProjectQuantityTracker.findUnique({
      where: {
        projectNo_stream_fontColor: {
          projectNo: item.projectNo,
          stream: item.stream,
          fontColor: item.fontColor,
        },
      },
    });

    if (!existing) {
      await prisma.planningProjectQuantityTracker.create({
        data: {
          projectNo: item.projectNo,
          stream: item.stream,
          fontColor: item.fontColor,
          projectShortname: item.pShort || null,
          projectName: item.pName || null,
          lastQuantity: item.totalQty,
          lastChangedDate: todayStr,
        },
      });
    } else if (existing.lastQuantity !== item.totalQty) {
      await prisma.planningProjectQuantityTracker.update({
        where: { id: existing.id },
        data: {
          lastQuantity: item.totalQty,
          lastChangedDate: todayStr,
          projectShortname: item.pShort || existing.projectShortname,
          projectName: item.pName || existing.projectName,
        },
      });
    }
  }

  // --------------------------------------------------------------------------
  // 2. PRODUCTION SERIES HISTORY
  // --------------------------------------------------------------------------
  const productionRows = datasetMap[RoleCode.PRODUCTION] || [];

  for (const row of productionRows) {
    const pShortRaw =
      findCellValue(row.data, 'Short Name') ||
      findCellValue(row.data, 'Project Shortname') ||
      findCellValue(row.data, 'Project Shortname (from planning column B)') ||
      findCellValue(row.data, 'Customer & Project Name') ||
      findCellValue(row.data, 'Project Name') ||
      row.rawCells?.[1] ||
      row.rawCells?.[0] ||
      row.rawCells?.[2];

    const pShort = cleanStr(pShortRaw);
    const pNo = cleanStr(findCellValue(row.data, 'Project No') || findCellValue(row.data, 'Project No.') || row.rawCells?.[0]);
    const pName = String(pShortRaw || '');

    let stream = '1';
    for (const sh of STREAM_HEADER_CANDIDATES) {
      const v = findCellValue(row.data, sh);
      if (v !== null && v !== undefined && v !== '') {
        stream = normalizeStream(v);
        break;
      }
    }

    const rawSeries =
      findCellValue(row.data, 'Series (from planning column G)') ||
      findCellValue(row.data, 'Series') ||
      findCellValue(row.data, 'Series No') ||
      row.rawCells?.[6];

    let seriesNumber = rawSeries ? parseInt(String(rawSeries).replace(/[^0-9]/g, ''), 10) : 0;
    if (!seriesNumber || isNaN(seriesNumber)) {
      seriesNumber = detectedProdSeries > 0 ? detectedProdSeries : 1;
    }

    const fontColor = normalizeColor(row.fontColor);

    const rawProduced =
      row.rawCells?.[16] ??
      findCellValue(row.data, '__COLUMN_Q__') ??
      findCellValue(row.data, 'Total Produced') ??
      findCellValue(row.data, 'Total Produced Quantity') ??
      findCellValue(row.data, 'Produced Quantity') ??
      findCellValue(row.data, 'Column Q') ??
      findCellValue(row.data, 'Produced (m2)');

    const totalProduced = parseNumeric(rawProduced);

    if (pShort && seriesNumber > 0 && totalProduced > 0) {
      await prisma.productionSeriesHistory.upsert({
        where: {
          projectShortname_stream_fontColor_seriesNumber: {
            projectShortname: pShort,
            stream,
            fontColor,
            seriesNumber,
          },
        },
        update: {
          totalProduced,
          projectNo: pNo || undefined,
          projectName: pName || undefined,
        },
        create: {
          projectShortname: pShort,
          projectNo: pNo || null,
          projectName: pName || null,
          stream,
          fontColor,
          seriesNumber,
          totalProduced,
        },
      });
    }
  }

  // --------------------------------------------------------------------------
  // 3. FETCH HISTORICAL PLANNING & PREVIOUS RUN FOR DISPATCH TRACKING
  // --------------------------------------------------------------------------
  const allQuantityTrackers = await prisma.planningProjectQuantityTracker.findMany();
  const allHistoricalPlanningSeries = await prisma.planningSeriesHistory.findMany({
    orderBy: { seriesNumber: 'asc' },
  });

  const previousRun = await prisma.mr11Run.findFirst({
    where: {
      status: 'READY',
      recordCount: { gt: 0 },
    },
    orderBy: { generatedAt: 'desc' },
  });

  const prevDispatchHistory: Record<string, { quantity: number; date: string }> = {
    ...((previousRun?.sourceSnapshot as any)?.dispatchTracker || {}),
  };

  if (previousRun && Array.isArray(previousRun.records)) {
    for (const r of previousRun.records as Record<string, any>[]) {
      const rShort = cleanStr(r['Short Name'] || r['Project Shortname'] || r['Project Short Code'] || r['Short Code'] || getProjectIdentifier(r));
      const rName = cleanStr(r['Customer & Project Name'] || r['Project Name']);
      const rStream = normalizeStream(r['Stream']);
      const rFont = normalizeColor(r['_fontColor']);
      const rFill = normalizeFillColor(r['_fillColor']);

      const k = `${rShort}__${rName}__${rStream}__${rFont}__${rFill}`;

      const q = parseNumeric(
        r['Total Dispatch'] ??
        r['Total Dispatched'] ??
        r['Total Dispatched Quantity'] ??
        r['Total Dispatch (m2)'] ??
        r['Total Dispatched (m2)']
      );

      let d: string | null = null;
      for (const key of Object.keys(r)) {
        const lk = key.toLowerCase();
        if (lk.includes('dispatch') && lk.includes('date') && r[key]) {
          d = String(r[key]).trim();
          break;
        }
      }

      if (q > 0 && !prevDispatchHistory[k]) {
        prevDispatchHistory[k] = { quantity: q, date: d || todayStr };
      }
    }
  }

  const newDispatchTracker: Record<string, { quantity: number; date: string }> = {};

  const bdRecords = datasetMap[RoleCode.BD];
  if (!bdRecords || bdRecords.length === 0) {
    const run = await prisma.mr11Run.create({
      data: {
        status: 'PARTIAL' as any,
        sourceSnapshot,
        recordCount: 0,
        records: [],
      },
    });
    return run.id;
  }

  // --------------------------------------------------------------------------
  // 4. DERIVE MASTER MR11 ROWS
  // --------------------------------------------------------------------------
  const designRows = datasetMap[RoleCode.DESIGN] || [];
  const shellplanRows = datasetMap[RoleCode.SHELLPLAN] || [];
  const dispatchRows = datasetMap[RoleCode.DISPATCH] || [];

  const derivedMr11Rows = bdRecords.map((bdRowItem) => {
    const bdData = bdRowItem.data;
    const projectNo = cleanStr(findCellValue(bdData, 'Project No') || findCellValue(bdData, 'Project No.') || bdRowItem.rawCells?.[1]);
    const shortName = cleanStr(findCellValue(bdData, 'Short Name') || findCellValue(bdData, 'Project Shortname') || bdRowItem.rawCells?.[2]);
    const projectName = cleanStr(findCellValue(bdData, 'Customer & Project Name') || findCellValue(bdData, 'Project Name') || bdRowItem.rawCells?.[0]);

    let bdStream = '1';
    for (const sh of STREAM_HEADER_CANDIDATES) {
      const v = findCellValue(bdData, sh);
      if (v !== null && v !== undefined && v !== '') {
        bdStream = normalizeStream(v);
        break;
      }
    }

    const bdFontColor = normalizeColor(bdRowItem.fontColor);
    const bdFillColor = normalizeFillColor(bdRowItem.fillColor);
    const outRow: Record<string, any> = {};
    const cellColors: Record<string, string> = {};

    // 1. Populate all defined BD/Pre-Shellplan & Commercial Columns directly from BD
    for (const mapping of MR11_ORDERED_COLUMNS) {
      if (mapping.sourceDept === RoleCode.BD) {
        outRow[mapping.target] = findCellValue(bdData, mapping.sourceColumn);
      }
    }

    // 2. Generic Department Fallbacks (for any finance/other keys)
    for (const mapping of MR11_ORDERED_COLUMNS) {
      if (
        mapping.sourceDept !== RoleCode.BD &&
        mapping.sourceDept !== RoleCode.DESIGN &&
        mapping.sourceDept !== RoleCode.SHELLPLAN &&
        mapping.sourceDept !== RoleCode.PLANNING &&
        mapping.sourceDept !== RoleCode.PRODUCTION &&
        mapping.sourceDept !== RoleCode.DISPATCH
      ) {
        const deptDataset = datasetMap[mapping.sourceDept] || [];
        const possibleKeyNames = MR11_SOURCE_KEY_MAP[mapping.sourceDept] || [];

        let bestCandidate: Record<string, any> | null = null;
        let highestScore = -1;

        for (const candidate of deptDataset) {
          const cData = candidate.data;
          let idMatched = false;
          for (const keyName of possibleKeyNames) {
            const raw = cleanStr(findCellValue(cData, keyName));
            if (raw && (raw === projectNo || raw === shortName || (shortName && raw.includes(shortName)))) {
              idMatched = true;
              break;
            }
          }
          if (!idMatched) continue;

          let score = 1;
          let cStream = '1';
          for (const sh of STREAM_HEADER_CANDIDATES) {
            const v = findCellValue(cData, sh);
            if (v !== null && v !== undefined && v !== '') {
              cStream = normalizeStream(v);
              break;
            }
          }
          if (cStream === bdStream) score += 4;

          const cFont = normalizeColor(candidate.fontColor);
          if (cFont === bdFontColor) score += 8;

          const cFill = normalizeFillColor(candidate.fillColor);
          if (bdFillColor && cFill && cFill === bdFillColor) score += 10;

          if (score > highestScore) {
            highestScore = score;
            bestCandidate = cData;
          }
        }

        outRow[mapping.target] = bestCandidate ? findCellValue(bestCandidate, mapping.sourceColumn) : null;
      }
    }

    // ------------------------------------------------------------------------
    // DISPATCH MAPPINGS: 5-POINT COMPOSITE KEY MATCHING
    // (Short Code + Project Name + Stream + Font Colour + Row Fill Colour)
    // ------------------------------------------------------------------------
    const matchedDispatchRows = dispatchRows.filter((dRow) => {
      const dShort = cleanStr(
        findCellValue(dRow.data, 'Short Name') ||
        findCellValue(dRow.data, 'Project Shortname') ||
        findCellValue(dRow.data, 'Project Shortname (from bd column C)') ||
        findCellValue(dRow.data, 'Project Short Code') ||
        findCellValue(dRow.data, 'Short Code') ||
        dRow.rawCells?.[1] ||
        dRow.rawCells?.[2]
      );

      const dName = cleanStr(
        findCellValue(dRow.data, 'Customer & Project Name') ||
        findCellValue(dRow.data, 'Project Name') ||
        findCellValue(dRow.data, 'Project Name (from bd column A)') ||
        dRow.rawCells?.[0] ||
        dRow.rawCells?.[1]
      );

      const dNo = cleanStr(
        findCellValue(dRow.data, 'Project No') ||
        findCellValue(dRow.data, 'Project No.') ||
        dRow.rawCells?.[0]
      );

      const shortMatched =
        (shortName && dShort && (dShort === shortName || dShort.includes(shortName) || shortName.includes(dShort))) ||
        (projectNo && dNo && (dNo === projectNo || dNo.includes(projectNo) || projectNo.includes(dNo)));

      if (!shortMatched) return false;

      if (projectName && dName && !(dName === projectName || dName.includes(projectName) || projectName.includes(dName))) {
        return false;
      }

      let dStream = '1';
      for (const sh of STREAM_HEADER_CANDIDATES) {
        const v = findCellValue(dRow.data, sh);
        if (v !== null && v !== undefined && v !== '') {
          dStream = normalizeStream(v);
          break;
        }
      }
      if (dStream !== bdStream) return false;

      const dFontColor = normalizeColor(dRow.fontColor);
      if (dFontColor !== bdFontColor) return false;

      const dFillColor = normalizeFillColor(dRow.fillColor);
      if (bdFillColor !== dFillColor) return false;

      return true;
    });

    // ------------------------------------------------------------------------
    // CUMULATIVE FORMWORK QUANTITY SAILED (m2)
    // ------------------------------------------------------------------------
    let totalFormworkSailed = 0;
    let hasSailedValue = false;

    for (const dRow of matchedDispatchRows) {
      const rawSailed =
        findCellValue(dRow.data, 'Formwork Quantity Sailed (m2)') ??
        findCellValue(dRow.data, 'Formwork Quantity Sailed m2') ??
        findCellValue(dRow.data, 'Formwork Quantity Sailed') ??
        findCellValue(dRow.data, 'Quantity Sailed (m2)') ??
        findCellValue(dRow.data, 'Quantity Sailed') ??
        findCellValue(dRow.data, 'Total Sailed (m2)') ??
        findCellValue(dRow.data, 'Total Sailed');

      const num = parseNumeric(rawSailed);
      if (!isNaN(num) && num > 0) {
        totalFormworkSailed += num;
        hasSailedValue = true;
      }
    }

    const finalFormworkSailed = hasSailedValue ? totalFormworkSailed : null;

    outRow['Formwork Quantity Sailed (m2)'] = finalFormworkSailed;
    outRow['Formwork Quantity Sailed m2'] = finalFormworkSailed;
    outRow['Formwork Quantity Sailed'] = finalFormworkSailed;

    if (Array.isArray(ORDERED_HEADER_LIST)) {
      ORDERED_HEADER_LIST.forEach((h) => {
        const cleanH = h.toLowerCase().trim();
        if (
          cleanH === 'formwork quantity sailed (m2)' ||
          cleanH === 'formwork quantity sailed m2' ||
          cleanH === 'formwork quantity sailed' ||
          cleanH === 'quantity sailed (m2)' ||
          (cleanH.includes('formwork') && cleanH.includes('sailed'))
        ) {
          outRow[h] = finalFormworkSailed;
        }
      });
    }

    // ------------------------------------------------------------------------
    // TOTAL DISPATCH (COLUMN K) & DISPATCHED DATE TRACKER
    // ------------------------------------------------------------------------
    let matchedDispatchRowForK: ExtractedRow | null = null;
    let directColumnKValue = 0;

    for (const dRow of matchedDispatchRows) {
      const val = parseNumeric(
        dRow.rawCells?.[10] ??
        findCellValue(dRow.data, '__COLUMN_K__') ??
        findCellValue(dRow.data, 'Column K') ??
        findCellValue(dRow.data, 'Total Dispatch') ??
        findCellValue(dRow.data, 'Total Dispatched') ??
        findCellValue(dRow.data, 'Total Dispatched Quantity') ??
        findCellValue(dRow.data, 'Dispatched Quantity') ??
        findCellValue(dRow.data, 'Total Dispatch (m2)') ??
        findCellValue(dRow.data, 'Total Dispatched (m2)')
      );
      if (val > directColumnKValue) {
        directColumnKValue = val;
        matchedDispatchRowForK = dRow;
      }
    }

    if (!matchedDispatchRowForK && matchedDispatchRows.length > 0) {
      matchedDispatchRowForK = matchedDispatchRows[matchedDispatchRows.length - 1];
    }

    const finalTotalDispatch = directColumnKValue > 0 ? directColumnKValue : null;

    outRow['Total Dispatch'] = finalTotalDispatch;
    outRow['Total Dispatched'] = finalTotalDispatch;
    outRow['Total Dispatched Quantity'] = finalTotalDispatch;
    outRow['Total Dispatch (m2)'] = finalTotalDispatch;
    outRow['Total Dispatched (m2)'] = finalTotalDispatch;
    outRow['Total Dispatched Quantity m2'] = finalTotalDispatch;

    if (Array.isArray(ORDERED_HEADER_LIST)) {
      ORDERED_HEADER_LIST.forEach((h) => {
        const cleanH = h.toLowerCase().trim();
        if (
          cleanH === 'total dispatch' ||
          cleanH === 'total dispatched' ||
          cleanH === 'total dispatched quantity' ||
          cleanH === 'total dispatch (m2)' ||
          cleanH === 'total dispatched (m2)' ||
          cleanH.includes('total dispatch') ||
          cleanH.includes('total dispatched')
        ) {
          outRow[h] = finalTotalDispatch;
        }
      });
    }

    const dispatchCompositeKey = `${shortName || projectNo}__${projectName}__${bdStream}__${bdFontColor}__${bdFillColor}`;
    const previousEntry = prevDispatchHistory[dispatchCompositeKey];
    let resolvedDispatchedDate: string | null = null;

    if (directColumnKValue > 0) {
      if (!previousEntry) {
        resolvedDispatchedDate = todayStr;
        newDispatchTracker[dispatchCompositeKey] = { quantity: directColumnKValue, date: todayStr };
      } else if (previousEntry.quantity !== directColumnKValue) {
        resolvedDispatchedDate = todayStr;
        newDispatchTracker[dispatchCompositeKey] = { quantity: directColumnKValue, date: todayStr };
      } else {
        resolvedDispatchedDate = previousEntry.date || todayStr;
        newDispatchTracker[dispatchCompositeKey] = { quantity: directColumnKValue, date: resolvedDispatchedDate };
      }
    } else {
      resolvedDispatchedDate = null;
    }

    outRow['Dispatched Date'] = resolvedDispatchedDate;
    outRow['Dispatch Date'] = resolvedDispatchedDate;
    outRow['Actual Dispatched Date'] = resolvedDispatchedDate;
    outRow['Actual Dispatch Date'] = resolvedDispatchedDate;
    outRow['Date Dispatched'] = resolvedDispatchedDate;

    if (Array.isArray(ORDERED_HEADER_LIST)) {
      ORDERED_HEADER_LIST.forEach((h) => {
        const cleanH = h.toLowerCase().trim();
        if (
          cleanH === 'dispatched date' ||
          cleanH === 'dispatch date' ||
          cleanH === 'actual dispatched date' ||
          cleanH === 'actual dispatch date' ||
          cleanH === 'date dispatched' ||
          (cleanH.includes('dispatch') && cleanH.includes('date'))
        ) {
          outRow[h] = resolvedDispatchedDate;
        }
      });
    }

    // ------------------------------------------------------------------------
    // ROBUST ATD LOGIC: COLUMN W (INDEX 22) VS COLUMNS P TO V (INDICES 15-21)
    // ------------------------------------------------------------------------
    let latestDateW: Date | null = null;
    let latestDateWStr: string | null = null;
    let latestDatePV: Date | null = null;
    let latestDatePVStr: string | null = null;

    for (const dRow of matchedDispatchRows) {
      const candidateWValues: any[] = [
        dRow.rawCells?.[22],
        findCellValue(dRow.data, '__COLUMN_W__'),
        dispatchHeaders[22] ? dRow.data[dispatchHeaders[22]] : undefined,
        findCellValue(dRow.data, 'Column W'),
        findCellValue(dRow.data, 'ATD'),
        findCellValue(dRow.data, 'ATD Date'),
        findCellValue(dRow.data, 'Actual Time of Departure'),
      ];

      for (const valW of candidateWValues) {
        if (valW !== undefined && valW !== null && String(valW).trim() !== '') {
          const d = parseFlexibleDate(valW);
          if (d) {
            if (!latestDateW || d.getTime() > latestDateW.getTime()) {
              latestDateW = d;
              latestDateWStr = formatDateString(d);
            }
          } else {
            const s = String(valW).trim();
            if (s && s !== '-' && !latestDateWStr) {
              latestDateWStr = s;
            }
          }
        }
      }

      for (let c = 15; c <= 21; c++) {
        const candidatePVValues: any[] = [
          dRow.rawCells?.[c],
          dispatchHeaders[c] ? dRow.data[dispatchHeaders[c]] : undefined,
        ];

        for (const valPV of candidatePVValues) {
          if (valPV !== undefined && valPV !== null && String(valPV).trim() !== '') {
            const d = parseFlexibleDate(valPV);
            if (d) {
              if (!latestDatePV || d.getTime() > latestDatePV.getTime()) {
                latestDatePV = d;
                latestDatePVStr = formatDateString(d);
              }
            } else {
              const s = String(valPV).trim();
              if (s && s !== '-' && !latestDatePVStr) {
                latestDatePVStr = s;
              }
            }
          }
        }
      }
    }

    let finalAtdDate: string | null = null;
    let atdColor: string = '#FFFFFF';

    if (latestDateWStr) {
      finalAtdDate = latestDateWStr;
      atdColor = '#FFFFFF';
    } else if (latestDatePVStr) {
      finalAtdDate = latestDatePVStr;
      atdColor = '#FFFF00';
    }

    outRow['ATD'] = finalAtdDate;
    outRow['ATD Date'] = finalAtdDate;
    outRow['Actual Time of Departure'] = finalAtdDate;
    outRow['_atdColor'] = atdColor;

    if (Array.isArray(ORDERED_HEADER_LIST)) {
      ORDERED_HEADER_LIST.forEach((h) => {
        const cleanH = h.toLowerCase().trim();
        if (
          cleanH === 'atd' ||
          cleanH === 'atd date' ||
          cleanH === 'actual time of departure' ||
          cleanH === 'actual departure date' ||
          cleanH.includes('atd')
        ) {
          outRow[h] = finalAtdDate;
          cellColors[h] = atdColor;
        }
      });
    }

    cellColors['ATD'] = atdColor;
    cellColors['ATD Date'] = atdColor;
    cellColors['Actual Time of Departure'] = atdColor;
    outRow['_cellColors'] = cellColors;

    // ------------------------------------------------------------------------
    // MONTHLY BREAKDOWN & ANNUAL TOTALS (2026 & 2027) MAPPINGS FROM DISPATCH
    // ------------------------------------------------------------------------
    const MONTH_COLUMNS_26 = ['Jan-26', 'Feb-26', 'Mar-26', 'Apr-26', 'May-26', 'Jun-26', 'Jul-26', 'Aug-26', 'Sep-26', 'Oct-26', 'Nov-26', 'Dec-26'];
    const MONTH_COLUMNS_27 = ['Jan-27', 'Feb-27', 'Mar-27', 'Apr-27', 'May-27', 'Jun-27', 'Jul-27', 'Aug-27', 'Sep-27', 'Oct-27', 'Nov-27', 'Dec-27'];

    let sum2026 = 0;
    let sum2027 = 0;

    MONTH_COLUMNS_26.forEach((mCol) => {
      let val = 0;
      for (const dRow of matchedDispatchRows) {
        const v = parseNumeric(findCellValue(dRow.data, mCol));
        if (v > 0) val += v;
      }
      outRow[mCol] = val > 0 ? val : null;
      sum2026 += val;
    });
    outRow['Total 2026 m2'] = sum2026 > 0 ? sum2026 : null;

    MONTH_COLUMNS_27.forEach((mCol) => {
      let val = 0;
      for (const dRow of matchedDispatchRows) {
        const v = parseNumeric(findCellValue(dRow.data, mCol));
        if (v > 0) val += v;
      }
      outRow[mCol] = val > 0 ? val : null;
      sum2027 += val;
    });
    outRow['Total 2027 m2'] = sum2027 > 0 ? sum2027 : null;

    // Fallback for remaining primary Dispatch non-volume attributes
    const primaryDispatchRow = matchedDispatchRowForK || (matchedDispatchRows.length > 0 ? matchedDispatchRows[0] : null);
    if (primaryDispatchRow) {
      for (const mapping of MR11_ORDERED_COLUMNS) {
        if (mapping.sourceDept === RoleCode.DISPATCH && !outRow[mapping.target]) {
          outRow[mapping.target] = findCellValue(primaryDispatchRow.data, mapping.sourceColumn);
        }
      }
    }

    // ------------------------------------------------------------------------
    // DESIGN MAPPINGS (COL AL, AM, AN): STRICTLY BY STREAM (COLOR IGNORED)
    // ------------------------------------------------------------------------
    const streamMatchedDesign = designRows.filter((dRow) => {
      const dProj = getProjectIdentifier(dRow.data);
      let dStream = '1';
      for (const sh of STREAM_HEADER_CANDIDATES) {
        const v = findCellValue(dRow.data, sh);
        if (v !== null && v !== undefined && v !== '') {
          dStream = normalizeStream(v);
          break;
        }
      }

      const projMatches =
        (projectNo && (dProj === projectNo || dProj.includes(projectNo) || projectNo.includes(dProj))) ||
        (shortName && (dProj === shortName || dProj.includes(shortName) || shortName.includes(dProj))) ||
        (projectName && (dProj === projectName || dProj.includes(projectName) || projectName.includes(dProj)));

      return projMatches && dStream === bdStream;
    });

    const designStatuses = streamMatchedDesign
      .map((d) => findCellValue(d.data, 'Formwork Design Status') || findCellValue(d.data, 'Design Status') || findCellValue(d.data, 'Status'))
      .filter((v) => v !== null && v !== undefined && v !== '');
    outRow['Formwork Design Status'] = resolveConsolidatedDesignStatus(designStatuses.map(String));

    const designDates = streamMatchedDesign
      .map((d) => findCellValue(d.data, 'Actual Formwork Order Completion Date') || findCellValue(d.data, 'Actual Completion Date') || findCellValue(d.data, 'Completion Date'))
      .filter((v) => v !== null && v !== undefined && v !== '');
    let latestDesignDate: string | null = null;
    for (const dv of designDates) {
      const fDate = formatDateString(dv);
      if (fDate && (!latestDesignDate || fDate > latestDesignDate)) {
        latestDesignDate = fDate;
      }
    }
    outRow['Actual Formwork Order Completion Date'] = latestDesignDate;

    let totalQuantityOrdered = 0;
    let hasQuantity = false;
    for (const d of streamMatchedDesign) {
      const rawQ =
        findCellValue(d.data, 'Total Quantity Ordered m2') ??
        findCellValue(d.data, 'Total Quantity Ordered (m2)') ??
        findCellValue(d.data, 'Total Quantity Ordered') ??
        findCellValue(d.data, 'Total Quantity') ??
        findCellValue(d.data, 'Quantity Ordered') ??
        findCellValue(d.data, 'Order Quantity');
      const num = parseNumeric(rawQ);
      if (!isNaN(num) && num > 0) {
        totalQuantityOrdered += num;
        hasQuantity = true;
      }
    }
    const finalQuantityAN = hasQuantity ? totalQuantityOrdered : null;
    outRow['Total Quantity Ordered m2'] = finalQuantityAN;
    outRow['Total Quantity Ordered (m2)'] = finalQuantityAN;
    outRow['Total Quantity Ordered'] = finalQuantityAN;

    if (Array.isArray(ORDERED_HEADER_LIST) && ORDERED_HEADER_LIST.length >= 40) {
      const colANHeader = ORDERED_HEADER_LIST[39];
      if (colANHeader) {
        outRow[colANHeader] = finalQuantityAN;
      }
    }

    // ------------------------------------------------------------------------
    // SHELLPLAN MAPPINGS (COL AJ, AK): STRICTLY BY STREAM (COLOR IGNORED)
    // ------------------------------------------------------------------------
    const streamMatchedShellplan = shellplanRows.filter((spRow) => {
      const spProj = getProjectIdentifier(spRow.data);
      let spStream = '1';
      for (const sh of STREAM_HEADER_CANDIDATES) {
        const v = findCellValue(spRow.data, sh);
        if (v !== null && v !== undefined && v !== '') {
          spStream = normalizeStream(v);
          break;
        }
      }
      const projMatches =
        (projectNo && (spProj === projectNo || spProj.includes(projectNo) || projectNo.includes(spProj))) ||
        (shortName && (spProj === shortName || spProj.includes(shortName) || shortName.includes(spProj))) ||
        (projectName && (spProj === projectName || spProj.includes(projectName) || projectName.includes(spProj)));
      return projMatches && spStream === bdStream;
    });

    const spStatuses = streamMatchedShellplan
      .map((c) => findCellValue(c.data, 'Shell Plan Status - Pending Consultant Drawings') || findCellValue(c.data, 'Shell Plan Status'))
      .filter((v) => v !== null && v !== undefined && v !== '');
    outRow['Shell Plan Status - Pending Consultant Drawings'] = resolveConsolidatedShellplanStatus(spStatuses.map(String));

    const spDates = streamMatchedShellplan
      .map((c) => findCellValue(c.data, 'Shell Plan Approved Date') || findCellValue(c.data, 'Approved Date') || findCellValue(c.data, 'Latest Submission Date'))
      .filter((v) => v !== null && v !== undefined && v !== '');
    let latestSpDate: string | null = null;
    for (const sv of spDates) {
      const fDate = formatDateString(sv);
      if (fDate && (!latestSpDate || fDate > latestSpDate)) {
        latestSpDate = fDate;
      }
    }
    outRow['Shell Plan Approved Date'] = latestSpDate;

    // ------------------------------------------------------------------------
    // PLANNING MAPPINGS (COL AO, AP): COLOR & ROW FILL AWARE
    // ------------------------------------------------------------------------
    const matchedPlanningSeries = allHistoricalPlanningSeries.filter((s) => {
      const pClean = cleanStr(s.projectNo);
      const sClean = cleanStr(s.projectShortname);
      const matchesId = pClean === projectNo || sClean === shortName || (shortName && sClean.includes(shortName));
      const matchesStream = s.stream === bdStream || s.stream === '1' || bdStream === '1';
      const sFont = normalizeColor(s.fontColor);
      return matchesId && matchesStream && sFont === bdFontColor;
    });

    if (matchedPlanningSeries.length > 0) {
      const sumProcessed = matchedPlanningSeries.reduce((acc, curr) => acc + (curr.totalProcessed || 0), 0);
      outRow['Total Processed'] = sumProcessed;
      outRow['Total Processed (m2)'] = sumProcessed;
    }

    const tracker = allQuantityTrackers.find(
      (t) =>
        (cleanStr(t.projectNo) === projectNo || (shortName && cleanStr(t.projectShortname) === shortName)) &&
        (t.stream === bdStream || t.stream === '1' || bdStream === '1') &&
        normalizeColor(t.fontColor) === bdFontColor
    );

    let matchedPlanningRow: ExtractedRow | null = null;
    let highestPlanScore = -1;

    for (const p of planningRows) {
      const pNo = cleanStr(findCellValue(p.data, 'Project No') || findCellValue(p.data, 'Project No.') || findCellValue(p.data, 'Project No. (from design column A)'));
      const pShort = cleanStr(findCellValue(p.data, 'Short Name') || findCellValue(p.data, 'Project Shortname') || findCellValue(p.data, 'Project Shortname (from bd column C)'));
      const idMatches = (projectNo && pNo === projectNo) || (shortName && pShort === shortName);
      if (!idMatches) continue;

      let score = 1;
      let pStream = '1';
      for (const sh of STREAM_HEADER_CANDIDATES) {
        const v = findCellValue(p.data, sh);
        if (v !== null && v !== undefined && v !== '') {
          pStream = normalizeStream(v);
          break;
        }
      }
      if (pStream === bdStream) score += 4;

      const pColor = normalizeColor(p.fontColor);
      const pFill = normalizeFillColor(p.fillColor);
      if (pColor === bdFontColor) score += 8;
      if (bdFillColor && pFill && pFill === bdFillColor) score += 10;

      if (score > highestPlanScore) {
        highestPlanScore = score;
        matchedPlanningRow = p;
      }
    }

    const activeClosingDate = matchedPlanningRow
      ? findCellValue(matchedPlanningRow.data, 'Closing Date ') ||
        findCellValue(matchedPlanningRow.data, 'Closing Date') ||
        findCellValue(matchedPlanningRow.data, 'Processed Date')
      : null;

    const latestSeriesClosingDate = matchedPlanningSeries.length > 0
      ? matchedPlanningSeries[matchedPlanningSeries.length - 1].closingDate
      : null;

    const resolvedProcessedDate =
      activeClosingDate ||
      latestSeriesClosingDate ||
      (tracker ? tracker.lastChangedDate : null) ||
      todayStr;

    outRow['Processed Date'] = resolvedProcessedDate;
    outRow['Closing Date'] = resolvedProcessedDate;

    if (Array.isArray(ORDERED_HEADER_LIST) && ORDERED_HEADER_LIST.length >= 42) {
      const colAPHeader = ORDERED_HEADER_LIST[41];
      if (colAPHeader) {
        outRow[colAPHeader] = resolvedProcessedDate;
      }
    }

    // ------------------------------------------------------------------------
    // PRODUCTION MAPPINGS (COL AQ, AR): COLOR & ROW FILL AWARE
    // ------------------------------------------------------------------------
    let matchedProdRow: ExtractedRow | null = null;
    let highestProdScore = -1;

    for (const pRow of productionRows) {
      const pShort = cleanStr(
        findCellValue(pRow.data, 'Short Name') ||
        findCellValue(pRow.data, 'Project Shortname') ||
        findCellValue(pRow.data, 'Project Shortname (from planning column B)') ||
        pRow.rawCells?.[1] ||
        pRow.rawCells?.[0] ||
        pRow.rawCells?.[2]
      );
      const pNo = cleanStr(
        findCellValue(pRow.data, 'Project No') ||
        findCellValue(pRow.data, 'Project No.') ||
        pRow.rawCells?.[0]
      );

      const idMatches =
        (shortName && pShort === shortName) ||
        (projectNo && pNo === projectNo) ||
        (shortName && pShort.includes(shortName)) ||
        (projectNo && pShort.includes(projectNo)) ||
        (projectName && pShort.includes(projectName));

      if (!idMatches) continue;

      let score = 1;
      let stream = '1';
      for (const sh of STREAM_HEADER_CANDIDATES) {
        const v = findCellValue(pRow.data, sh);
        if (v !== null && v !== undefined && v !== '') {
          stream = normalizeStream(v);
          break;
        }
      }
      if (stream === bdStream) score += 4;

      const pColor = normalizeColor(pRow.fontColor);
      const pFill = normalizeFillColor(pRow.fillColor);

      if (pColor === bdFontColor) score += 8;
      if (bdFillColor && pFill && pFill === bdFillColor) score += 10;

      if (score > highestProdScore) {
        highestProdScore = score;
        matchedProdRow = pRow;
      }
    }

    const directColumnQValue = matchedProdRow
      ? parseNumeric(
          matchedProdRow.rawCells?.[16] ??
          findCellValue(matchedProdRow.data, '__COLUMN_Q__') ??
          findCellValue(matchedProdRow.data, 'Total Produced') ??
          findCellValue(matchedProdRow.data, 'Total Produced Quantity') ??
          findCellValue(matchedProdRow.data, 'Produced Quantity') ??
          findCellValue(matchedProdRow.data, 'Column Q') ??
          findCellValue(matchedProdRow.data, 'Produced (m2)')
        )
      : 0;

    const finalColumnAQ = directColumnQValue > 0 ? directColumnQValue : null;

    outRow['Total Produced'] = finalColumnAQ;
    outRow['Total Produced Quantity'] = finalColumnAQ;

    if (Array.isArray(ORDERED_HEADER_LIST) && ORDERED_HEADER_LIST.length >= 43) {
      const colAQHeader = ORDERED_HEADER_LIST[42];
      if (colAQHeader) {
        outRow[colAQHeader] = finalColumnAQ;
      }
    }

    let latestFilledDate: string | null = null;

    if (matchedProdRow) {
      const COL_R_INDEX = 17;
      const COL_AV_INDEX = 47;

      for (let c = COL_AV_INDEX; c >= COL_R_INDEX; c--) {
        const cellVal = matchedProdRow.rawCells?.[c];
        const headerName = productionHeaders[c];
        const valFromHeader = headerName ? matchedProdRow.data[headerName] : undefined;
        const targetVal = cellVal !== undefined ? cellVal : valFromHeader;

        if (isCellFilled(targetVal)) {
          if (headerName) {
            const formatted = formatDateString(headerName);
            if (formatted) {
              latestFilledDate = formatted;
              break;
            }
          }

          const cellAsDate = formatDateString(targetVal);
          if (cellAsDate) {
            latestFilledDate = cellAsDate;
            break;
          }

          if (headerName) {
            latestFilledDate = headerName;
            break;
          }
        }
      }

      if (!latestFilledDate) {
        const fallbackDate = findCellValue(matchedProdRow.data, 'Day/Date') || findCellValue(matchedProdRow.data, 'Date');
        latestFilledDate = formatDateString(fallbackDate);
      }
    }

    outRow['Produced Date'] = latestFilledDate;
    if (Array.isArray(ORDERED_HEADER_LIST) && ORDERED_HEADER_LIST.length >= 44) {
      const colARHeader = ORDERED_HEADER_LIST[43];
      if (colARHeader) {
        outRow[colARHeader] = latestFilledDate;
      }
    }

    outRow['_fontColor'] = bdFontColor;
    outRow['_fillColor'] = bdFillColor;

    return outRow;
  });

  // --------------------------------------------------------------------------
  // 5. STRICT SORTING: PROJECT -> STREAM -> FONT COLOR (BLACK FIRST) -> ROW COLOR
  // --------------------------------------------------------------------------
  derivedMr11Rows.sort((a, b) => {
    const projA = getProjectIdentifier(a);
    const projB = getProjectIdentifier(b);
    if (projA !== projB) {
      return projA.localeCompare(projB);
    }

    const streamA = parseFloat(String(a['Stream'] || '1').replace(/[^0-9.]/g, '')) || 1;
    const streamB = parseFloat(String(b['Stream'] || '1').replace(/[^0-9.]/g, '')) || 1;
    if (streamA !== streamB) {
      return streamA - streamB;
    }

    const colorA = normalizeColor(a['_fontColor']);
    const colorB = normalizeColor(b['_fontColor']);

    const isBlackA = colorA === '#000000';
    const isBlackB = colorB === '#000000';

    if (isBlackA && !isBlackB) return -1;
    if (!isBlackA && isBlackB) return 1;
    if (colorA !== colorB) return colorA.localeCompare(colorB);

    const fillA = normalizeFillColor(a['_fillColor']);
    const fillB = normalizeFillColor(b['_fillColor']);
    return fillA.localeCompare(fillB);
  });

  // --------------------------------------------------------------------------
  // 6. ATTACH STREAM MERGE METADATA (FOR DESIGN & SHELLPLAN UI MERGING)
  // --------------------------------------------------------------------------
  for (let i = 0; i < derivedMr11Rows.length; ) {
    const curProj = getProjectIdentifier(derivedMr11Rows[i]);
    const curStream = normalizeStream(derivedMr11Rows[i]['Stream']);
    let span = 1;

    while (
      i + span < derivedMr11Rows.length &&
      getProjectIdentifier(derivedMr11Rows[i + span]) === curProj &&
      normalizeStream(derivedMr11Rows[i + span]['Stream']) === curStream
    ) {
      span++;
    }

    derivedMr11Rows[i]['_isStreamLead'] = true;
    derivedMr11Rows[i]['_streamSpan'] = span;

    for (let j = 1; j < span; j++) {
      derivedMr11Rows[i + j]['_isStreamLead'] = false;
      derivedMr11Rows[i + j]['_streamSpan'] = 0;
    }

    i += span;
  }

  const run = await prisma.mr11Run.create({
    data: {
      status: 'READY' as any,
      sourceSnapshot: {
        ...sourceSnapshot,
        dispatchTracker: newDispatchTracker,
      },
      recordCount: derivedMr11Rows.length,
      records: derivedMr11Rows as any,
    },
  });

  return run.id;
}