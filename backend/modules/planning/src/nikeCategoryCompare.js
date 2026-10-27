const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

// Month dictionary for fast lookup
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_MAP = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

function normalizeHeaderStr(val) {
  return String(val || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/**
 * Universal UTC Date Parser
 * Handles Excel serial numbers, JS Dates, DD/MM/YYYY, YYYY-MM-DD, 8-Apr-2026, etc.
 */
function parseDateUtc(val) {
  if (val === null || val === undefined || val === '') return null;

  // If already Date object
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return {
      year: val.getUTCFullYear(),
      month: val.getUTCMonth(), // 0-11
      day: val.getUTCDate(),
    };
  }

  // If Excel serial number (e.g. 45000)
  if (typeof val === 'number') {
    if (val <= 0 || isNaN(val)) return null;
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return {
        year: parsed.y,
        month: parsed.m - 1, // 0-11
        day: parsed.d,
      };
    }
    return null;
  }

  const str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'n/a' || str.toLowerCase() === 'null') return null;

  // If number formatted as string
  if (/^\d+(\.\d+)?$/.test(str)) {
    const num = parseFloat(str);
    if (num > 30000 && num < 70000) {
      const parsed = XLSX.SSF.parse_date_code(num);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        return { year: parsed.y, month: parsed.m - 1, day: parsed.d };
      }
    }
  }

  // Formats: DD-MM-YYYY, DD/MM/YYYY, DD.MM.YYYY or YYYY-MM-DD
  const parts = str.split(/[/.-]/);
  if (parts.length === 3) {
    let year, month, day;
    if (parts[0].length === 4) {
      // YYYY-MM-DD
      year = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      day = parseInt(parts[2], 10);
    } else if (isNaN(parts[1])) {
      // e.g. 8-Apr-2026 or 23-Sep-26
      day = parseInt(parts[0], 10);
      const mKey = parts[1].toLowerCase().slice(0, 3);
      month = MONTH_MAP[mKey] !== undefined ? MONTH_MAP[mKey] : -1;
      year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;
    } else {
      // Ambiguous DD/MM/YYYY vs MM/DD/YYYY - Indian standard is DD/MM/YYYY
      day = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;
    }

    if (!isNaN(year) && !isNaN(month) && month >= 0 && month < 12 && !isNaN(day) && day >= 1 && day <= 31) {
      return { year, month, day };
    }
  }

  // Fallback: Date.parse
  const ts = Date.parse(str);
  if (!isNaN(ts)) {
    const d = new Date(ts);
    return {
      year: d.getUTCFullYear(),
      month: d.getUTCMonth(),
      day: d.getUTCDate(),
    };
  }

  return null;
}

/**
 * Parses raw cell value from ExcelJS cell, resolving formula cached results
 */
function getCellCalculatedValue(cell) {
  if (!cell || cell.value === null || cell.value === undefined) return null;
  if (typeof cell.value === 'object' && cell.value !== null) {
    if (cell.value instanceof Date) return cell.value;
    if ('result' in cell.value) {
      return cell.value.result;
    }
  }
  return cell.value;
}

/**
 * Parse a Nike Snapshot file buffer into structured rows
 */
async function parseNikeSnapshot(fileBuffer, fileName = 'Snapshot') {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error(`Workbook "${fileName}" has no sheets.`);
  }

  // Look for "Summary sheet" (case-insensitive) or fallback to first sheet
  let worksheet = workbook.worksheets.find(
    (ws) => ws.name && ws.name.trim().toLowerCase() === 'summary sheet'
  );
  const usedFallbackSheet = !worksheet;
  if (!worksheet) {
    worksheet = workbook.worksheets[0];
  }

  const warnings = [];
  if (usedFallbackSheet) {
    warnings.push(`Sheet "Summary sheet" not found in "${fileName}". Using first sheet "${worksheet.name}".`);
  }

  // Scan first 10 rows for headers
  let headerRowIdx = -1;
  let colIndices = {
    category: -1,
    exFactory: -1,
    qty: -1,
    status: -1,
    gac: -1,
    style: -1,
  };

  const totalRowCount = worksheet.rowCount || 0;
  if (totalRowCount === 0) {
    throw new Error(`Sheet "${worksheet.name}" in "${fileName}" is empty.`);
  }

  for (let r = 1; r <= Math.min(10, totalRowCount); r++) {
    const row = worksheet.getRow(r);
    let catIdx = -1, exIdx = -1, qtyIdx = -1, stIdx = -1, gacIdx = -1, styIdx = -1;

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const rawHeader = String(cell.value || '').trim();
      const norm = normalizeHeaderStr(rawHeader);

      if (['category', 'categ'].includes(norm)) {
        if (catIdx === -1) catIdx = colNumber;
      } else if (['ex factory', 'ex-factory', 'exfactory', 'ex factory date', 'ex factory '].includes(norm)) {
        if (exIdx === -1) exIdx = colNumber;
      } else if (['qty', 'quantity', 'qty '].includes(norm)) {
        if (qtyIdx === -1) {
          qtyIdx = colNumber;
        } else {
          warnings.push(`File "${fileName}" has duplicate Qty header at column ${colNumber}. Using first column.`);
        }
      } else if (['status', 'status '].includes(norm)) {
        if (stIdx === -1) stIdx = colNumber;
      } else if (['gac', 'gac date'].includes(norm)) {
        if (gacIdx === -1) gacIdx = colNumber;
      } else if (['style', 'styles'].includes(norm)) {
        if (styIdx === -1) styIdx = colNumber;
      }
    });

    if (catIdx !== -1 && exIdx !== -1 && qtyIdx !== -1) {
      headerRowIdx = r;
      colIndices = {
        category: catIdx,
        exFactory: exIdx,
        qty: qtyIdx,
        status: stIdx,
        gac: gacIdx,
        style: styIdx,
      };
      break;
    }
  }

  if (headerRowIdx === -1) {
    const missing = [];
    if (colIndices.category === -1) missing.push('Category');
    if (colIndices.exFactory === -1) missing.push('EX FACTORY');
    if (colIndices.qty === -1) missing.push('Qty');
    if (colIndices.status === -1) missing.push('Status');
    throw new Error(
      `Could not locate required headers in "${fileName}". Missing or unrecognized columns: ${missing.join(', ')}.`
    );
  }

  if (colIndices.status === -1) {
    warnings.push(`File "${fileName}" is missing a "Status" column. All rows will be counted under ALL only.`);
  }

  const validRows = [];
  let unassignedCount = 0;
  let unassignedQty = 0;
  let nonNumericQtyCount = 0;
  let negativeQtyCount = 0;
  let rawTotalQty = 0;
  const unknownStatusCounts = {};

  // Process rows from headerRowIdx + 1 to end
  for (let r = headerRowIdx + 1; r <= totalRowCount; r++) {
    const row = worksheet.getRow(r);
    
    // Read category, exFactory, qty, status
    const rawCategory = row.getCell(colIndices.category).value;
    const rawExFactoryCell = row.getCell(colIndices.exFactory);
    const rawGacCell = colIndices.gac !== -1 ? row.getCell(colIndices.gac) : null;
    const rawQty = row.getCell(colIndices.qty).value;
    const rawStatus = colIndices.status !== -1 ? row.getCell(colIndices.status).value : '';
    const rawStyle = colIndices.style !== -1 ? row.getCell(colIndices.style).value : '';

    const categoryStr = String(rawCategory || '').trim().toUpperCase();
    const styleStr = String(rawStyle || '').trim();
    const statusStr = String(rawStatus || '').trim().toLowerCase();

    // Check if completely empty row
    if (!rawCategory && !rawExFactoryCell.value && !rawQty && !rawStatus && !rawStyle) {
      continue;
    }

    // Check Grand Total / Total row
    if (
      categoryStr === 'TOTAL' ||
      categoryStr === 'GRAND TOTAL' ||
      categoryStr === 'SUBTOTAL' ||
      categoryStr.includes('GRAND TOTAL')
    ) {
      continue; // Skip total row
    }

    // Parse Quantity
    let qty = 0;
    if (rawQty !== null && rawQty !== undefined && rawQty !== '') {
      if (typeof rawQty === 'number') {
        qty = isNaN(rawQty) ? 0 : rawQty;
      } else {
        const cleanedStr = String(rawQty).replace(/,/g, '').trim();
        const parsedNum = parseFloat(cleanedStr);
        if (isNaN(parsedNum)) {
          qty = 0;
          nonNumericQtyCount++;
        } else {
          qty = parsedNum;
        }
      }
    }

    if (qty < 0) {
      negativeQtyCount++;
    }

    rawTotalQty += qty;

    // Parse Date: try EX FACTORY calculated value first
    let exValue = getCellCalculatedValue(rawExFactoryCell);
    let parsedDate = parseDateUtc(exValue);

    // Fallback if formula had no cached result and GAC is available
    if (!parsedDate && rawGacCell) {
      const gacVal = getCellCalculatedValue(rawGacCell);
      const parsedGac = parseDateUtc(gacVal);
      if (parsedGac) {
        // GAC minus 4 days
        const gacDate = new Date(Date.UTC(parsedGac.year, parsedGac.month, parsedGac.day));
        gacDate.setUTCDate(gacDate.getUTCDate() - 4);
        parsedDate = {
          year: gacDate.getUTCFullYear(),
          month: gacDate.getUTCMonth(),
          day: gacDate.getUTCDate(),
        };
        warnings.push(`Row ${r}: EX FACTORY formula had no cached result; fell back to GAC - 4 days.`);
      }
    }

    if (!parsedDate) {
      if (qty !== 0) {
        unassignedCount++;
        unassignedQty += qty;
      }
      continue;
    }

    // Status classification
    let normalizedStatus = 'other';
    if (statusStr === 'confirmed') {
      normalizedStatus = 'confirmed';
    } else if (statusStr === 'projections' || statusStr === 'projection') {
      normalizedStatus = 'projections';
    } else {
      normalizedStatus = 'other';
      if (statusStr) {
        if (!unknownStatusCounts[statusStr]) {
          unknownStatusCounts[statusStr] = { count: 0, qty: 0 };
        }
        unknownStatusCounts[statusStr].count++;
        unknownStatusCounts[statusStr].qty += qty;
      }
    }

    const monthKey = `${parsedDate.year}-${String(parsedDate.month + 1).padStart(2, '0')}`;

    validRows.push({
      category: categoryStr || 'UNCLASSIFIED',
      year: parsedDate.year,
      month: parsedDate.month, // 0-11
      monthKey,
      qty,
      status: normalizedStatus,
      rawStatus: statusStr,
      style: styleStr,
    });
  }

  if (unassignedCount > 0) {
    warnings.push(`${unassignedCount} rows (${unassignedQty.toLocaleString('en-IN')} qty) in "${fileName}" have no valid EX FACTORY date.`);
  }
  if (nonNumericQtyCount > 0) {
    warnings.push(`${nonNumericQtyCount} rows in "${fileName}" had non-numeric Qty values (treated as 0).`);
  }
  if (negativeQtyCount > 0) {
    warnings.push(`${negativeQtyCount} rows in "${fileName}" had negative Qty values.`);
  }

  for (const [statusVal, info] of Object.entries(unknownStatusCounts)) {
    warnings.push(
      `File "${fileName}" contains unclassified status "${statusVal}" (${info.count} rows, ${info.qty.toLocaleString('en-IN')} qty) counted in ALL only.`
    );
  }

  return {
    fileName,
    rows: validRows,
    rawTotalQty,
    unassignedQty,
    warnings,
  };
}

/**
 * Format month label according to the reference and rules:
 * Months in the same year as the reference year show plain (Oct, Nov); other years show suffix (Jan'27).
 */
function getMonthDisplayLabel(monthKey, referenceYear) {
  const [yearStr, monthStr] = monthKey.split('-');
  const y = parseInt(yearStr, 10);
  const m = parseInt(monthStr, 10) - 1; // 0-11
  const mName = MONTH_NAMES[m] || 'M' + monthStr;

  if (referenceYear && y === referenceYear) {
    return mName;
  }
  const shortYear = String(y).slice(-2);
  return `${mName}'${shortYear}`;
}

/**
 * Generate all union months and compute default From/To
 */
function computeAvailableMonths(prevData, currData) {
  const monthMap = new Map();

  function record(monthKey) {
    if (!monthMap.has(monthKey)) {
      const [yearStr, monthStr] = monthKey.split('-');
      monthMap.set(monthKey, {
        key: monthKey,
        year: parseInt(yearStr, 10),
        month: parseInt(monthStr, 10) - 1,
        prevQty: 0,
        currQty: 0,
      });
    }
  }

  for (const r of prevData.rows) {
    record(r.monthKey);
    monthMap.get(r.monthKey).prevQty += r.qty;
  }
  for (const r of currData.rows) {
    record(r.monthKey);
    monthMap.get(r.monthKey).currQty += r.qty;
  }

  const sortedMonths = Array.from(monthMap.values()).sort((a, b) => {
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });

  if (sortedMonths.length === 0) {
    return {
      months: [],
      defaultFrom: '',
      defaultTo: '',
    };
  }

  // Find first month where total qty differs
  let defaultFrom = sortedMonths[0].key;
  for (const m of sortedMonths) {
    if (Math.round(m.prevQty) !== Math.round(m.currQty)) {
      defaultFrom = m.key;
      break;
    }
  }

  // Default To = From + 5 months (clamped)
  const fromIdx = sortedMonths.findIndex((m) => m.key === defaultFrom);
  const toIdx = Math.min(sortedMonths.length - 1, fromIdx + 5);
  const defaultTo = sortedMonths[toIdx].key;

  return {
    months: sortedMonths.map((m) => ({
      key: m.key,
      year: m.year,
      month: m.month,
      label: `${MONTH_NAMES[m.month]} ${m.year}`,
    })),
    defaultFrom,
    defaultTo,
  };
}

/**
 * Core business comparison function shared between Preview and Export
 */
function buildComparison(prevData, currData, fromMonthKey, toMonthKey, prevLabel = 'Previous', currLabel = 'Current') {
  // Validate and sort month range
  const allMonthKeysSet = new Set();
  for (const r of prevData.rows) allMonthKeysSet.add(r.monthKey);
  for (const r of currData.rows) allMonthKeysSet.add(r.monthKey);

  let sortedAllMonths = Array.from(allMonthKeysSet).sort();

  if (sortedAllMonths.length === 0) {
    throw new Error('No valid monthly data found in either file.');
  }

  // Fallbacks if not provided
  if (!fromMonthKey) fromMonthKey = sortedAllMonths[0];
  if (!toMonthKey) toMonthKey = sortedAllMonths[sortedAllMonths.length - 1];

  if (fromMonthKey > toMonthKey) {
    throw new Error(`"From Month" (${fromMonthKey}) cannot be later than "To Month" (${toMonthKey}).`);
  }

  // Determine active months in range [fromMonthKey, toMonthKey]
  // Note: Generate full contiguous month sequence or union of months
  const [fromY, fromM] = fromMonthKey.split('-').map(Number);
  const [toY, toM] = toMonthKey.split('-').map(Number);

  const activeMonths = [];
  let curY = fromY;
  let curM = fromM; // 1-12
  while (curY < toY || (curY === toY && curM <= toM)) {
    const k = `${curY}-${String(curM).padStart(2, '0')}`;
    activeMonths.push({
      key: k,
      year: curY,
      month: curM - 1,
    });
    curM++;
    if (curM > 12) {
      curM = 1;
      curY++;
    }
  }

  const referenceYear = fromY;
  const monthHeaders = activeMonths.map((m) => ({
    key: m.key,
    label: getMonthDisplayLabel(m.key, referenceYear),
  }));

  // Determine Categories
  // Standard categories: BOTTOM, JACKET, TOP first
  const standardCats = ['BOTTOM', 'JACKET', 'TOP'];
  const otherCatsSet = new Set();

  for (const r of prevData.rows) {
    if (!standardCats.includes(r.category)) {
      otherCatsSet.add(r.category);
    }
  }
  for (const r of currData.rows) {
    if (!standardCats.includes(r.category)) {
      otherCatsSet.add(r.category);
    }
  }

  const sortedOtherCats = Array.from(otherCatsSet).sort();
  const categories = [...standardCats, ...sortedOtherCats];

  // Helper to aggregate data for a specific status filter
  // statusFilter: 'all' | 'confirmed' | 'projections'
  function aggregateTable(statusFilter) {
    const tableRows = [];
    const totals = {
      category: 'Total',
      months: {},
    };

    for (const m of activeMonths) {
      totals.months[m.key] = { prev: 0, curr: 0, diff: 0 };
    }

    for (const cat of categories) {
      const row = {
        category: cat,
        months: {},
      };

      for (const m of activeMonths) {
        // Prev sum
        let prevSum = 0;
        for (const r of prevData.rows) {
          if (r.monthKey === m.key && r.category === cat) {
            if (statusFilter === 'all' || r.status === statusFilter) {
              prevSum += r.qty;
            }
          }
        }

        // Curr sum
        let currSum = 0;
        for (const r of currData.rows) {
          if (r.monthKey === m.key && r.category === cat) {
            if (statusFilter === 'all' || r.status === statusFilter) {
              currSum += r.qty;
            }
          }
        }

        const diff = currSum - prevSum;

        row.months[m.key] = {
          prev: prevSum,
          curr: currSum,
          diff: diff,
        };

        totals.months[m.key].prev += prevSum;
        totals.months[m.key].curr += currSum;
        totals.months[m.key].diff += diff;
      }

      tableRows.push(row);
    }

    return {
      rows: tableRows,
      totals,
    };
  }

  const allTable = aggregateTable('all');
  const confirmedTable = aggregateTable('confirmed');
  const projectionsTable = aggregateTable('projections');

  // Title strings
  const fromLabel = getMonthDisplayLabel(fromMonthKey, referenceYear);
  const toLabel = getMonthDisplayLabel(toMonthKey, referenceYear);
  const rangeTitleSuffix = `${fromLabel} to ${toLabel}`;

  const tables = [
    {
      id: 'all',
      title: `ALL (Confirmed + Projections) : ${rangeTitleSuffix}`,
      ...allTable,
    },
    {
      id: 'confirmed',
      title: `CONFIRMED ONLY : ${rangeTitleSuffix}`,
      ...confirmedTable,
    },
    {
      id: 'projections',
      title: `PROJECTIONS ONLY : ${rangeTitleSuffix}`,
      ...projectionsTable,
    },
  ];

  // Perform Sanity Checks
  const sanityErrors = [];

  // Check 1: Confirmed + Projections + Other === All for every cell
  for (let cIdx = 0; cIdx < categories.length; cIdx++) {
    const cat = categories[cIdx];
    for (const m of activeMonths) {
      const allP = allTable.rows[cIdx].months[m.key].prev;
      const confP = confirmedTable.rows[cIdx].months[m.key].prev;
      const projP = projectionsTable.rows[cIdx].months[m.key].prev;

      const allC = allTable.rows[cIdx].months[m.key].curr;
      const confC = confirmedTable.rows[cIdx].months[m.key].curr;
      const projC = projectionsTable.rows[cIdx].months[m.key].curr;

      // Find any "other" status sum
      let otherP = 0;
      for (const r of prevData.rows) {
        if (r.monthKey === m.key && r.category === cat && r.status === 'other') otherP += r.qty;
      }
      let otherC = 0;
      for (const r of currData.rows) {
        if (r.monthKey === m.key && r.category === cat && r.status === 'other') otherC += r.qty;
      }

      if (Math.round(confP + projP + otherP) !== Math.round(allP)) {
        sanityErrors.push(`Discrepancy in Previous [${cat} - ${m.key}]: Confirmed (${confP}) + Proj (${projP}) != All (${allP}).`);
      }
      if (Math.round(confC + projC + otherC) !== Math.round(allC)) {
        sanityErrors.push(`Discrepancy in Current [${cat} - ${m.key}]: Confirmed (${confC}) + Proj (${projC}) != All (${allC}).`);
      }
    }
  }

  // Check 2: Sum of ALL months + unassigned qty === sum of raw valid Qty column
  let prevAllMonthsSum = 0;
  for (const r of prevData.rows) prevAllMonthsSum += r.qty;
  const prevCheckDiff = Math.abs(prevAllMonthsSum + prevData.unassignedQty - prevData.rawTotalQty);
  if (prevCheckDiff > 0.01) {
    sanityErrors.push(
      `Previous File Total Mismatch: sum of months (${prevAllMonthsSum}) + unassigned (${prevData.unassignedQty}) != raw total (${prevData.rawTotalQty}).`
    );
  }

  let currAllMonthsSum = 0;
  for (const r of currData.rows) currAllMonthsSum += r.qty;
  const currCheckDiff = Math.abs(currAllMonthsSum + currData.unassignedQty - currData.rawTotalQty);
  if (currCheckDiff > 0.01) {
    sanityErrors.push(
      `Current File Total Mismatch: sum of months (${currAllMonthsSum}) + unassigned (${currData.unassignedQty}) != raw total (${currData.rawTotalQty}).`
    );
  }

  const allWarnings = [
    ...(prevData.warnings || []),
    ...(currData.warnings || []),
  ];

  return {
    monthHeaders,
    activeMonths,
    categories,
    tables,
    prevLabel,
    currLabel,
    fromMonthKey,
    toMonthKey,
    fromLabel,
    toLabel,
    grandTotals: {
      prev: {
        total: prevAllMonthsSum,
        confirmed: prevData.rows.filter((r) => r.status === 'confirmed').reduce((s, r) => s + r.qty, 0),
        projections: prevData.rows.filter((r) => r.status === 'projections').reduce((s, r) => s + r.qty, 0),
      },
      curr: {
        total: currAllMonthsSum,
        confirmed: currData.rows.filter((r) => r.status === 'confirmed').reduce((s, r) => s + r.qty, 0),
        projections: currData.rows.filter((r) => r.status === 'projections').reduce((s, r) => s + r.qty, 0),
      },
    },
    sanityChecks: {
      passed: sanityErrors.length === 0,
      errors: sanityErrors,
    },
    warnings: allWarnings,
  };
}

/**
 * Generate Excel workbook matching formatting requirements
 */
async function generateCategoryComparisonExcel(comparisonResult) {
  const { monthHeaders, categories, tables, prevLabel, currLabel, warnings } = comparisonResult;

  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Category Comparison', {
    views: [{ showGridLines: true }],
  });

  // Setup Column Widths
  // Col A = 14, each month has 3 columns = 12 each
  const cols = [{ width: 14 }];
  for (let i = 0; i < monthHeaders.length * 3; i++) {
    cols.push({ width: 12 });
  }
  ws.columns = cols;

  const fontName = 'Times New Roman';
  const standardFont = { name: fontName, size: 11, color: { argb: 'FF000000' } };
  const boldFont = { name: fontName, size: 11, bold: true, color: { argb: 'FF000000' } };
  const titleFont = { name: fontName, size: 11, bold: true, color: { argb: 'FF000000' } };

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };

  const lightBlueFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
  const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  const salmonFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8CBAD' } };
  const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
  const redFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };

  const greenFont = { name: fontName, size: 11, color: { argb: 'FF006100' } };
  const redFont = { name: fontName, size: 11, color: { argb: 'FF9C0006' } };
  const boldGreenFont = { name: fontName, size: 11, bold: true, color: { argb: 'FF006100' } };
  const boldRedFont = { name: fontName, size: 11, bold: true, color: { argb: 'FF9C0006' } };

  // Indian number format
  const indianNumFmt = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0';

  let currentRow = 1;
  const totalTableCols = 1 + monthHeaders.length * 3;

  for (let tIdx = 0; tIdx < tables.length; tIdx++) {
    const table = tables[tIdx];

    // Table Title
    const titleCell = ws.getCell(currentRow, 1);
    titleCell.value = table.title;
    titleCell.font = titleFont;
    ws.mergeCells(currentRow, 1, currentRow, totalTableCols);
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    currentRow++;

    const headerRow1Num = currentRow;
    const headerRow2Num = currentRow + 1;

    // Header Col A "Category" (merged 2 rows)
    ws.mergeCells(headerRow1Num, 1, headerRow2Num, 1);
    const catHeaderCell = ws.getCell(headerRow1Num, 1);
    catHeaderCell.value = 'Category';
    catHeaderCell.font = boldFont;
    catHeaderCell.fill = lightBlueFill;
    catHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    catHeaderCell.border = thinBorder;
    ws.getCell(headerRow2Num, 1).border = thinBorder;

    // Month headers (merged 3 columns each) and sub-headers
    for (let mIdx = 0; mIdx < monthHeaders.length; mIdx++) {
      const startCol = 2 + mIdx * 3;
      const endCol = startCol + 2;

      // Month name
      ws.mergeCells(headerRow1Num, startCol, headerRow1Num, endCol);
      const mCell = ws.getCell(headerRow1Num, startCol);
      mCell.value = monthHeaders[mIdx].label;
      mCell.font = boldFont;
      mCell.fill = yellowFill;
      mCell.alignment = { horizontal: 'center', vertical: 'middle' };

      for (let c = startCol; c <= endCol; c++) {
        ws.getCell(headerRow1Num, c).border = thinBorder;
      }

      // Sub-headers: Prev, Curr, DIFF
      const prevSubCell = ws.getCell(headerRow2Num, startCol);
      prevSubCell.value = prevLabel;
      prevSubCell.font = boldFont;
      prevSubCell.fill = yellowFill;
      prevSubCell.alignment = { horizontal: 'center', vertical: 'middle' };
      prevSubCell.border = thinBorder;

      const currSubCell = ws.getCell(headerRow2Num, startCol + 1);
      currSubCell.value = currLabel;
      currSubCell.font = boldFont;
      currSubCell.fill = yellowFill;
      currSubCell.alignment = { horizontal: 'center', vertical: 'middle' };
      currSubCell.border = thinBorder;

      const diffSubCell = ws.getCell(headerRow2Num, startCol + 2);
      diffSubCell.value = 'DIFF';
      diffSubCell.font = boldFont;
      diffSubCell.fill = salmonFill;
      diffSubCell.alignment = { horizontal: 'center', vertical: 'middle' };
      diffSubCell.border = thinBorder;
    }

    currentRow = headerRow2Num + 1;
    const firstDataRowNum = currentRow;

    // Data rows
    for (let rIdx = 0; rIdx < table.rows.length; rIdx++) {
      const rowData = table.rows[rIdx];
      const rNum = currentRow;

      const catCell = ws.getCell(rNum, 1);
      catCell.value = rowData.category;
      catCell.font = standardFont;
      catCell.border = thinBorder;
      catCell.alignment = { horizontal: 'left', vertical: 'middle' };

      for (let mIdx = 0; mIdx < monthHeaders.length; mIdx++) {
        const mKey = monthHeaders[mIdx].key;
        const mData = rowData.months[mKey] || { prev: 0, curr: 0, diff: 0 };
        const prevCol = 2 + mIdx * 3;
        const currCol = prevCol + 1;
        const diffCol = prevCol + 2;

        const prevColLetter = ws.getColumn(prevCol).letter;
        const currColLetter = ws.getColumn(currCol).letter;

        // Prev Cell
        const pCell = ws.getCell(rNum, prevCol);
        pCell.value = mData.prev;
        pCell.font = standardFont;
        pCell.border = thinBorder;
        if (mData.prev === 0) {
          pCell.numFmt = '"-"';
          pCell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else {
          pCell.numFmt = indianNumFmt;
          pCell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        // Curr Cell
        const cCell = ws.getCell(rNum, currCol);
        cCell.value = mData.curr;
        cCell.font = standardFont;
        cCell.border = thinBorder;
        if (mData.curr === 0) {
          cCell.numFmt = '"-"';
          cCell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else {
          cCell.numFmt = indianNumFmt;
          cCell.alignment = { horizontal: 'right', vertical: 'middle' };
        }

        // DIFF Cell: Formula =Current - Previous
        const dCell = ws.getCell(rNum, diffCol);
        dCell.value = {
          formula: `${currColLetter}${rNum}-${prevColLetter}${rNum}`,
          result: mData.diff,
        };
        dCell.border = thinBorder;

        if (mData.diff === 0) {
          dCell.font = standardFont;
          dCell.numFmt = '"-"';
          dCell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else if (mData.diff > 0) {
          dCell.fill = greenFill;
          dCell.font = greenFont;
          dCell.numFmt = indianNumFmt;
          dCell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          dCell.fill = redFill;
          dCell.font = redFont;
          dCell.numFmt = indianNumFmt;
          dCell.alignment = { horizontal: 'right', vertical: 'middle' };
        }
      }

      currentRow++;
    }

    const lastDataRowNum = currentRow - 1;
    const totalRowNum = currentRow;

    // Total Row
    const totalLabelCell = ws.getCell(totalRowNum, 1);
    totalLabelCell.value = 'Total';
    totalLabelCell.font = boldFont;
    totalLabelCell.border = thinBorder;
    totalLabelCell.alignment = { horizontal: 'left', vertical: 'middle' };

    for (let mIdx = 0; mIdx < monthHeaders.length; mIdx++) {
      const mKey = monthHeaders[mIdx].key;
      const mTotal = table.totals.months[mKey] || { prev: 0, curr: 0, diff: 0 };
      const prevCol = 2 + mIdx * 3;
      const currCol = prevCol + 1;
      const diffCol = prevCol + 2;

      const prevColLetter = ws.getColumn(prevCol).letter;
      const currColLetter = ws.getColumn(currCol).letter;

      // Prev Total Formula
      const pTotCell = ws.getCell(totalRowNum, prevCol);
      pTotCell.value = {
        formula: `SUM(${prevColLetter}${firstDataRowNum}:${prevColLetter}${lastDataRowNum})`,
        result: mTotal.prev,
      };
      pTotCell.font = boldFont;
      pTotCell.border = thinBorder;
      if (mTotal.prev === 0) {
        pTotCell.numFmt = '"-"';
        pTotCell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        pTotCell.numFmt = indianNumFmt;
        pTotCell.alignment = { horizontal: 'right', vertical: 'middle' };
      }

      // Curr Total Formula
      const cTotCell = ws.getCell(totalRowNum, currCol);
      cTotCell.value = {
        formula: `SUM(${currColLetter}${firstDataRowNum}:${currColLetter}${lastDataRowNum})`,
        result: mTotal.curr,
      };
      cTotCell.font = boldFont;
      cTotCell.border = thinBorder;
      if (mTotal.curr === 0) {
        cTotCell.numFmt = '"-"';
        cTotCell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        cTotCell.numFmt = indianNumFmt;
        cTotCell.alignment = { horizontal: 'right', vertical: 'middle' };
      }

      // DIFF Total Formula
      const dTotCell = ws.getCell(totalRowNum, diffCol);
      dTotCell.value = {
        formula: `${currColLetter}${totalRowNum}-${prevColLetter}${totalRowNum}`,
        result: mTotal.diff,
      };
      dTotCell.border = thinBorder;

      if (mTotal.diff === 0) {
        dTotCell.font = boldFont;
        dTotCell.numFmt = '"-"';
        dTotCell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (mTotal.diff > 0) {
        dTotCell.fill = greenFill;
        dTotCell.font = boldGreenFont;
        dTotCell.numFmt = indianNumFmt;
        dTotCell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else {
        dTotCell.fill = redFill;
        dTotCell.font = boldRedFont;
        dTotCell.numFmt = indianNumFmt;
        dTotCell.alignment = { horizontal: 'right', vertical: 'middle' };
      }
    }

    currentRow = totalRowNum + 3; // 2 blank rows between tables
  }

  // Notes block if warnings exist
  if (warnings && warnings.length > 0) {
    const notesTitleCell = ws.getCell(currentRow, 1);
    notesTitleCell.value = 'Notes / Audit Warnings:';
    notesTitleCell.font = boldFont;
    currentRow++;

    for (const w of warnings) {
      const noteCell = ws.getCell(currentRow, 1);
      noteCell.value = `• ${w}`;
      noteCell.font = standardFont;
      currentRow++;
    }
  }

  return await workbook.xlsx.writeBuffer();
}

module.exports = {
  parseNikeSnapshot,
  computeAvailableMonths,
  buildComparison,
  generateCategoryComparisonExcel,
  getMonthDisplayLabel,
};
