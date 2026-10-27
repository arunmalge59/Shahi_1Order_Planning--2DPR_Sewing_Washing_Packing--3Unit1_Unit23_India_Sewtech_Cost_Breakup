const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

const OUTPUT_HEADERS = [
  'O.CODE', '', '', '', '', '',
  'O.PROD', 'O.PROD', 'O.CUST', 'O^DD:1', 'O^DQ:1', 'O.CONTRACT_QTY',
  'O.UDSEASON', 'O.UDPCD', 'O.STATUS', 'O.effy', 'O.Descrip', 'O.UDBuy details',
  'O.Sprice', 'O.Udconversion', 'END',
];

const CP_HEADERS = [
  'Season', 'Unit', 'Item no', 'Style', 'Description', 'Colour', 'Order',
  'Projection qty', 'PO NO', 'OPD WK', 'PO qty', 'Total Qty', 'Proj pcd',
  'Proj xfty', 'Wash', 'Additional process details', 'Remarks', 'Plnd pcd',
  'Sew start', 'Sew end', 'Pos del', 'Extension',
];

const CUSTOMER = 'H&M';
const ORDER_STATUS = 'P';
const UD_CONVERSION = 75;

function resolveEffy(row) {
  return '';
}

function normalizedHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeText(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

const text = normalizeText;

function csvCell(value) {
  const cell = value === undefined || value === null ? '' : String(value);
  return /[",\r\n]/.test(cell) ? `"${cell.replace(/"/g, '""')}"` : cell;
}

function normalizeQuantity(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const valueText = text(value).replace(/,/g, '');
  if (!valueText) return null;
  const parsed = Number(valueText);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeOrderType(value) {
  const valueText = text(value).toLowerCase();
  if (!valueText) return '';
  if (valueText === 'sh') return 'SH';
  return valueText.replace(/\b[a-z]/g, (letter) => letter.toUpperCase());
}

function toExcelDateSerial(value) {
  const date = normalizeExcelDate(value);
  if (!date) return value;
  return Math.round((Date.UTC(date.year, date.month - 1, date.day) - Date.UTC(1899, 11, 30)) / 86400000);
}

function normalizeSeason(value) {
  return text(value).toUpperCase().replace(/\(.*?\)/g, '').replace(/\./g, '').replace(/\s+/g, '');
}

function normalizeWeekCode(value) {
  const match = text(value).match(/^wk-(\d{1,2})$/i);
  return match ? `wk-${match[1]}` : '';
}

function normalizeYear(value) {
  const year = Number(value);
  return year < 100 ? 2000 + year : year;
}

function monthNumber(value) {
  const months = {
    jan: 1, january: 1,
    feb: 2, february: 2,
    mar: 3, march: 3,
    apr: 4, april: 4,
    may: 5,
    jun: 6, june: 6,
    jul: 7, july: 7,
    aug: 8, august: 8,
    sep: 9, sept: 9, september: 9,
    oct: 10, october: 10,
    nov: 11, november: 11,
    dec: 12, december: 12,
  };
  return months[text(value).toLowerCase()] || null;
}

function normalizeExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value);
    return date ? { year: date.y, month: date.m, day: date.d } : null;
  }

  const valueText = text(value);
  if (!valueText) return null;

  let match = valueText.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2}|\d{4})$/);
  if (match) return { year: normalizeYear(match[3]), month: +match[2], day: +match[1] };

  match = valueText.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (match) return { year: +match[1], month: +match[2], day: +match[3] };

  match = valueText.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (match) return { year: +match[1], month: +match[2], day: +match[3] };

  match = valueText.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{2}|\d{4})$/);
  if (match) {
    const month = monthNumber(match[2]);
    return month ? { year: normalizeYear(match[3]), month, day: +match[1] } : null;
  }

  return null;
}

function formatOrderDate(value) {
  const date = normalizeExcelDate(value);
  return date ? `${String(date.day).padStart(2, '0')}-${String(date.month).padStart(2, '0')}-${date.year}` : '';
}

function formatCodeDate(value) {
  const date = normalizeExcelDate(value);
  return date ? `${date.year}${String(date.month).padStart(2, '0')}${String(date.day).padStart(2, '0')}` : '';
}

function getColumn(row, name) {
  const wanted = normalizedHeader(name);
  const key = Object.keys(row).find((candidate) => normalizedHeader(candidate) === wanted);
  return key === undefined ? '' : row[key];
}

function getColour(row) {
  return text(getColumn(row, 'Colour')) || text(getColumn(row, 'Color'));
}

function findHeaderRowIndex(matrix) {
  if (!Array.isArray(matrix) || !matrix.length) return 0;
  for (let i = 0; i < Math.min(matrix.length, 10); i++) {
    const row = matrix[i];
    if (Array.isArray(row)) {
      const hasHeader = row.some((cell) => {
        const h = normalizedHeader(cell);
        return h === 'itemno' || h === 'unit' || h === 'style' || h === 'season' || h === 'pono';
      });
      if (hasHeader) return i;
    }
  }
  return 0;
}

function parseMarketingFile(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: true });
  const marketingRows = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    if (matrix.length <= 1) continue;

    const headerRowIndex = findHeaderRowIndex(matrix);
    const headers = matrix[headerRowIndex].map((header) => text(header));
    const carryForwardColumns = new Set(['itemno', 'style', 'description']);
    const previousValues = {};
    const sheetSeason = normalizeSeason(sheetName);

    for (let i = headerRowIndex + 1; i < matrix.length; i++) {
      const values = matrix[i];
      const row = { sourceRowNumber: i + 1, sheetName, season: sheetSeason };

      headers.forEach((header, columnIndex) => {
        if (!header) return;
        const value = values[columnIndex] ?? '';
        const headerKey = normalizedHeader(header);
        row[header] = value === '' && carryForwardColumns.has(headerKey) && previousValues[headerKey] !== undefined
          ? previousValues[headerKey]
          : value;
        if (row[header] !== '') previousValues[headerKey] = row[header];
      });

      const hasData = Object.keys(row).some((key) => !['sourceRowNumber', 'sheetName', 'season'].includes(key) && text(row[key]));
      if (hasData) {
        row.__rawValues = values;
        marketingRows.push(row);
      }
    }
  }

  return marketingRows;
}

function getMarketingColumnByPosition(row, columnIndex) {
  return row.__rawValues?.[columnIndex] ?? '';
}

function skipCpTemplateRow(marketingRow, reason) {
  console.warn('Skipped H&M Non-Denim CP template row:', {
    sourceRowNumber: marketingRow.sourceRowNumber,
    sheetName: marketingRow.sheetName,
    reason,
    style: text(getColumn(marketingRow, 'Style')),
    season: marketingSeason(marketingRow),
    itemNo: text(getColumn(marketingRow, 'Item no')),
  });
}

function buildPreviousCpUnitLookup(previousCpBuffer) {
  if (!previousCpBuffer) return { byItemNo: new Map(), byItemAndPo: new Map(), itemHistory: new Map() };

  const { cpRows } = parseCpFile(previousCpBuffer);
  const unitsByItemNo = new Map();
  const unitsByItemAndPo = new Map();
  const itemHistory = new Map();

  for (const row of cpRows) {
    const itemNo = text(getColumn(row, 'Item no'));
    const poNo = text(getColumn(row, 'PO NO'));
    const unit = text(getColumn(row, 'Unit'));

    if (!itemNo) continue;

    if (unit) {
      // Track all units this item has used in the reference file
      if (!itemHistory.has(itemNo)) itemHistory.set(itemNo, new Set());
      itemHistory.get(itemNo).add(unit);

      // Match by Item No (if unique unit for item)
      if (!unitsByItemNo.has(itemNo)) unitsByItemNo.set(itemNo, new Set());
      unitsByItemNo.get(itemNo).add(unit);

      // Match by Item No + PO No
      if (poNo) {
        const itemPoKey = `${itemNo}|${poNo}`;
        if (!unitsByItemAndPo.has(itemPoKey)) unitsByItemAndPo.set(itemPoKey, new Set());
        unitsByItemAndPo.get(itemPoKey).add(unit);
      }
    }
  }

  return { byItemNo: unitsByItemNo, byItemAndPo: unitsByItemAndPo, itemHistory };
}

function unitFromPreviousCp(marketingRow, previousCpLookupData) {
  const itemNo = text(getColumn(marketingRow, 'Item no'));
  const poNo = text(getColumn(marketingRow, 'PO NO'));

  if (!itemNo) return { unit: '', flagType: null, comment: '' };

  const candidateUnits = previousCpLookupData.byItemNo?.get(itemNo);
  const historyUnits = previousCpLookupData.itemHistory?.get(itemNo);

  // 1. Single Unit by Item No
  if (candidateUnits && candidateUnits.size === 1) {
    return { unit: [...candidateUnits][0], flagType: null, comment: '' };
  }

  // 2. Exact match by Item No + PO No
  if (poNo) {
    const itemPoKey = `${itemNo}|${poNo}`;
    const poUnits = previousCpLookupData.byItemAndPo?.get(itemPoKey);
    if (poUnits && poUnits.size === 1) {
      return { unit: [...poUnits][0], flagType: null, comment: '' };
    }
  }

  // If item exists in reference history (has produced in the past in 1 or more units)
  if (historyUnits && historyUnits.size > 0) {
    const unitList = [...historyUnits].join(', ');
    const comment = `Needs manual Unit assignment. This Item+PO combo wasn't seen in reference week. Item '${itemNo}' has been produced in these Units before: ${unitList}. Pick the correct one for PO ${poNo || '(none)'}.`;
    return { unit: '', flagType: 'orange', comment };
  }

  // Else: Item has zero history in reference file at all (brand new item)
  const comment = `Needs manual Unit assignment. Item '${itemNo}' has no history in reference week at all (likely a brand-new style this week) — assign Unit manually.`;
  return { unit: '', flagType: 'yellow', comment };
}

async function generateHmNonDenimCpFile(marketingBuffer, previousCpBuffer = null) {
  const marketingRows = parseMarketingFile(marketingBuffer);
  if (!marketingRows.length) throw new Error('No data rows found in the Marketing file.');

  const previousCpLookupData = buildPreviousCpUnitLookup(previousCpBuffer);
  let generatedCpRows = 0;
  let autoFilled = 0;
  let flaggedItemKnown = 0;
  let flaggedNewItem = 0;
  const flaggedRows = [];
  const rowDataList = [];

  for (const marketingRow of marketingRows) {
    const poNo = getColumn(marketingRow, 'PO NO');
    const hasPoNo = text(poNo) !== '';
    const colour = getColumn(marketingRow, 'Color');
    const orderType = normalizeOrderType(getMarketingColumnByPosition(marketingRow, 6));
    const projectionQty = hasPoNo ? '' : getMarketingColumnByPosition(marketingRow, 7);
    const poQty = hasPoNo ? getColumn(marketingRow, 'PO qty') : '';
    const totalQty = hasPoNo ? poQty : projectionQty;
    const projectionQtyText = text(projectionQty);

    if (!text(colour) && !text(poNo) && !text(orderType)) {
      skipCpTemplateRow(marketingRow, 'Missing Colour, PO NO, and Order type');
      continue;
    }

    if (!text(projectionQty) && !text(poQty) && !text(poNo)) {
      skipCpTemplateRow(marketingRow, 'No Projection qty, PO qty, or PO NO found');
      continue;
    }

    if (!hasPoNo && projectionQtyText && normalizeQuantity(projectionQty) === null) {
      console.warn('H&M Non-Denim CP template warning: Marketing PROJ column has non-numeric value:', {
        sourceRowNumber: marketingRow.sourceRowNumber,
        sheetName: marketingRow.sheetName,
        value: projectionQtyText,
      });
    }

    // Get Unit from previous CP file (may be flagged orange or yellow)
    const unitResult = unitFromPreviousCp(marketingRow, previousCpLookupData);
    const unit = unitResult.unit;

    if (unitResult.flagType === 'orange') {
      flaggedItemKnown++;
      flaggedRows.push({
        sourceRowNumber: marketingRow.sourceRowNumber,
        sheetName: marketingRow.sheetName,
        style: text(getColumn(marketingRow, 'Style')),
        itemNo: text(getColumn(marketingRow, 'Item no')),
        flagType: 'orange',
        reason: unitResult.comment,
      });
    } else if (unitResult.flagType === 'yellow') {
      flaggedNewItem++;
      flaggedRows.push({
        sourceRowNumber: marketingRow.sourceRowNumber,
        sheetName: marketingRow.sheetName,
        style: text(getColumn(marketingRow, 'Style')),
        itemNo: text(getColumn(marketingRow, 'Item no')),
        flagType: 'yellow',
        reason: unitResult.comment,
      });
    } else if (unit) {
      autoFilled++;
    }

    const rowValues = [
      marketingSeason(marketingRow),
      unit,
      getColumn(marketingRow, 'Item no'),
      getColumn(marketingRow, 'Style'),
      getColumn(marketingRow, 'Description'),
      colour,
      orderType,
      projectionQty,
      poNo,
      getColumn(marketingRow, 'OPD WK'),
      poQty,
      totalQty,
      toExcelDateSerial(getColumn(marketingRow, 'Proj pcd')),
      toExcelDateSerial(getColumn(marketingRow, 'Proj xfty')),
      getColumn(marketingRow, 'Wash'),
      getColumn(marketingRow, 'Additional process details'),
      getColumn(marketingRow, 'Remarks'),
      '',
      '',
      '',
      '',
      '',
    ];

    rowDataList.push({
      values: rowValues,
      flagType: unitResult.flagType,
      comment: unitResult.comment,
    });

    generatedCpRows++;
  }

  // Create Excel workbook using ExcelJS to preserve fill colors and comments
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('H&M CP File');

  // Row 1: empty row
  worksheet.addRow([]);
  // Row 2: Header row
  const headerRow = worksheet.addRow(CP_HEADERS);
  headerRow.font = { bold: true };

  // Data rows (Row 3 onwards)
  for (const item of rowDataList) {
    const row = worksheet.addRow(item.values);

    // Unit cell is column 2 (B)
    const unitCell = row.getCell(2);
    if (item.flagType === 'orange') {
      unitCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC864' },
      };
      unitCell.note = item.comment;
    } else if (item.flagType === 'yellow') {
      unitCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' },
      };
      unitCell.note = item.comment;
    }

    // Date formatting for Proj pcd (col 13) and Proj xfty (col 14)
    for (const colIdx of [13, 14]) {
      const dateCell = row.getCell(colIdx);
      if (typeof dateCell.value === 'number') {
        dateCell.numFmt = 'mm-dd-yy';
      }
    }
  }

  worksheet.columns.forEach((column, index) => {
    const header = CP_HEADERS[index] || '';
    column.width = Math.max(14, header.length + 3);
  });

  const buffer = await workbook.xlsx.writeBuffer();

  console.log(`H&M Non-Denim CP File generated: ${generatedCpRows} total rows (${autoFilled} auto-filled, ${flaggedItemKnown} flagged orange - item known, ${flaggedNewItem} flagged yellow - brand new item)`);

  return {
    buffer: Buffer.from(buffer),
    summary: {
      totalMarketingRows: marketingRows.length,
      generatedCpRows,
      autoFilled,
      flaggedItemKnown,
      flaggedNewItem,
      flaggedForManualReview: flaggedItemKnown + flaggedNewItem,
    },
    flaggedRows,
  };
}
function parseCpFile(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('CP file has no sheets.');

  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true });
  if (!matrix.length) throw new Error('CP file is empty.');

  const headerRowIndex = findHeaderRowIndex(matrix);
  const headers = matrix[headerRowIndex].map((header) => text(header));
  const cpRows = [];

  for (let i = headerRowIndex + 1; i < matrix.length; i++) {
    const values = matrix[i];
    const row = { sourceRowNumber: i + 1 };
    headers.forEach((header, columnIndex) => {
      if (header) row[header] = values[columnIndex] ?? '';
    });
    if (Object.keys(row).some((key) => key !== 'sourceRowNumber' && text(row[key]))) cpRows.push(row);
  }

  return { sheetName, cpRows };
}

function marketingSeason(row) {
  return normalizeSeason(getColumn(row, 'Season')) || row.season || '';
}

function rowSeason(row) {
  return normalizeSeason(getColumn(row, 'Season'));
}

function buildJoinKeys(row, source) {
  const season = source === 'marketing' ? marketingSeason(row) : rowSeason(row);
  const itemNo = text(getColumn(row, 'Item no'));
  const poNo = text(getColumn(row, 'PO NO'));
  const style = text(getColumn(row, 'Style'));
  const colour = getColour(row);
  const opdWk = text(getColumn(row, 'OPD WK')).toLowerCase();

  return {
    primary: season && itemNo && poNo ? `${season}|${itemNo}|${poNo}` : '',
    fallback: season && style && colour && opdWk ? `${season}|${style}|${colour}|${opdWk}` : '',
  };
}

function addToMap(map, key, row) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(row);
}

function baseValidation(row, reason) {
  return {
    sourceRowNumber: row.sourceRowNumber,
    reason,
    style: text(getColumn(row, 'Style')),
    season: rowSeason(row),
    itemNo: text(getColumn(row, 'Item no')),
  };
}

function determineStatus(cpRow, validationErrors) {
  const projectionQtyRaw = getColumn(cpRow, 'Projection qty');
  const poQtyRaw = getColumn(cpRow, 'PO qty');
  const hasProjectionQty = text(projectionQtyRaw) !== '';
  const hasPoQty = text(poQtyRaw) !== '';

  if (hasProjectionQty && hasPoQty) {
    const validationError = baseValidation(cpRow, 'Both Projection qty and PO qty populated; defaulted to Confirmed');
    console.warn('H&M Non-Denim status conflict:', validationError);
    validationErrors.push(validationError);
    return 'Confirmed';
  }

  if (hasPoQty) return 'Confirmed';
  if (hasProjectionQty) return 'Projection';
  return '';
}

function getOrderQuantity(cpRow, status) {
  return status === 'Confirmed'
    ? normalizeQuantity(getColumn(cpRow, 'PO qty'))
    : normalizeQuantity(getColumn(cpRow, 'Projection qty'));
}

function buildOCode(style, season, weekCode, unit, status, exFactoryDate) {
  const codeDate = formatCodeDate(exFactoryDate);
  return status === 'Confirmed'
    ? `${style}-${season}::${weekCode}::${unit}::Confirmed::${codeDate}`
    : `${style}-${season}::${weekCode}::${unit}::${codeDate}`;
}

function skip(validationErrors, cpRow, reason) {
  const validationError = baseValidation(cpRow, reason);
  console.warn('Skipped invalid H&M Non-Denim row:', validationError);
  validationErrors.push(validationError);
}

function generateHmNonDenimOrderFile(marketingBuffer, cpBuffer, weekCode, previousCpBuffer = null) {
  const normalizedWeekCode = normalizeWeekCode(weekCode);
  if (!normalizedWeekCode) throw new Error('Week Code must be in the format wk-34.');

  const marketingRows = parseMarketingFile(marketingBuffer);
  const { sheetName, cpRows } = parseCpFile(cpBuffer);
  if (!marketingRows.length) throw new Error('No data rows found in the Marketing file.');
  if (!cpRows.length) throw new Error('No data rows found in the CP file.');

  const previousCpLookupData = buildPreviousCpUnitLookup(previousCpBuffer);

  const marketingByPrimary = new Map();
  const marketingByFallback = new Map();
  const marketingMatched = new Set();

  for (const row of marketingRows) {
    const keys = buildJoinKeys(row, 'marketing');
    addToMap(marketingByPrimary, keys.primary, row);
    addToMap(marketingByFallback, keys.fallback, row);
  }

  const validationErrors = [];
  const records = [];
  const matchedCpRows = new Set();
  const oCodeCounts = new Map();

  for (const cpRow of cpRows) {
    const cpKeys = buildJoinKeys(cpRow, 'cp');
    const marketingMatches = (cpKeys.primary && marketingByPrimary.get(cpKeys.primary)) || (cpKeys.fallback && marketingByFallback.get(cpKeys.fallback));
    const marketingRow = marketingMatches?.[0] || null;

    if (!marketingRow) {
      skip(validationErrors, cpRow, 'No matching Marketing row found');
      continue;
    }

    matchedCpRows.add(cpRow);
    marketingMatched.add(marketingRow);

    const status = determineStatus(cpRow, validationErrors);
    if (!status) {
      skip(validationErrors, cpRow, 'Neither Projection qty nor PO qty populated');
      continue;
    }

    const quantity = getOrderQuantity(cpRow, status);
    const totalQty = normalizeQuantity(getColumn(cpRow, 'Total Qty'));
    const style = text(getColumn(cpRow, 'Style'));
    const description = text(getColumn(cpRow, 'Description')) || style;
    const season = rowSeason(cpRow);

    // Get Unit from CP row. If blank, attempt auto-fill using previous week's CP file lookup
    let unit = text(getColumn(cpRow, 'Unit'));
    if (!unit && previousCpLookupData) {
      const unitResult = unitFromPreviousCp(cpRow, previousCpLookupData);
      if (unitResult.unit) {
        unit = unitResult.unit;
      }
    }

    const opdWeek = text(getColumn(cpRow, 'OPD WK'));
    const exFactoryDateValue = getColumn(cpRow, 'Proj xfty');
    const pcdDateValue = getColumn(cpRow, 'Proj pcd');
    const reasons = [];

    if (quantity === null) reasons.push(`Invalid ${status === 'Confirmed' ? 'PO qty' : 'Projection qty'}`);
    if (totalQty === null) reasons.push('Invalid Total Qty');
    if (!style) reasons.push('Missing Style');
    if (!season) reasons.push('Missing Season');
    if (!unit) reasons.push('Missing Unit');
    if (!opdWeek) reasons.push('Missing OPD WK');
    if (!formatCodeDate(exFactoryDateValue)) reasons.push('Missing or invalid Proj xfty');
    if (!formatOrderDate(pcdDateValue)) reasons.push('Missing or invalid Proj pcd');

    if (reasons.length) {
      skip(validationErrors, cpRow, reasons.join('; '));
      continue;
    }

    const baseOCode = buildOCode(style, season, normalizedWeekCode, unit, status, exFactoryDateValue);
    let oCode = baseOCode;

    if (!oCodeCounts.has(baseOCode)) {
      oCodeCounts.set(baseOCode, 1);
    } else {
      const count = oCodeCounts.get(baseOCode);
      oCodeCounts.set(baseOCode, count + 1);
      oCode = `${baseOCode}_${count}`;
    }

    const product = `${style}-${season}`;
    records.push([
      oCode,
      '::',
      normalizedWeekCode,
      status,
      unit,
      formatCodeDate(exFactoryDateValue),
      product,
      product,
      CUSTOMER,
      formatOrderDate(exFactoryDateValue),
      quantity,
      totalQty,
      season,
      formatOrderDate(pcdDateValue),
      ORDER_STATUS,
      resolveEffy(cpRow),
      description,
      '',
      '',
      UD_CONVERSION,
      'END',
    ]);
  }

  for (const row of marketingRows) {
    if (!marketingMatched.has(row)) {
      const validationError = {
        sourceRowNumber: row.sourceRowNumber,
        sheetName: row.sheetName,
        reason: 'No matching CP row found',
        style: text(getColumn(row, 'Style')),
        season: marketingSeason(row),
        itemNo: text(getColumn(row, 'Item no')),
      };
      console.warn('Unmatched H&M Non-Denim Marketing row:', validationError);
      validationErrors.push(validationError);
    }
  }

  const csvRows = records.map((record) => record.map(csvCell).join(','));
  const csv = [OUTPUT_HEADERS.map(csvCell).join(','), ...csvRows].join('\r\n') + '\r\n';

  return {
    csv,
    sheetName,
    summary: {
      totalCpRows: cpRows.length,
      totalMarketingRows: marketingRows.length,
      generatedOrders: records.length,
      unmatchedCpRows: cpRows.length - matchedCpRows.size,
      unmatchedMarketingRows: marketingRows.length - marketingMatched.size,
      validationErrors: validationErrors.length,
    },
    preview: records.slice(0, 100).map((record) => Object.fromEntries(OUTPUT_HEADERS.map((header, index) => [header || `Unnamed: ${index}`, record[index]]))),
    validationErrors,
  };
}

module.exports = {
  generateHmNonDenimCpFile,
  generateHmNonDenimOrderFile,
  normalizeText,
  normalizeQuantity,
  normalizeExcelDate,
  formatOrderDate,
  formatCodeDate,
  buildOCode,
  CP_HEADERS,
  OUTPUT_HEADERS,
};

