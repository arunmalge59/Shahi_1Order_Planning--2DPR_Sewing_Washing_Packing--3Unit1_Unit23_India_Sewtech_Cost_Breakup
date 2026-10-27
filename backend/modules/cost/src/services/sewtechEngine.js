/**
 * Sewtech Processing Engine
 *
 * Implements the core business logic for Sewtech Fashions Ltd. (B10 & B11 units):
 *  1. Workbook tab date scanning & range filtering
 *  2. Dynamic unit table & column detection (no hardcoded row/column coordinates)
 *  3. LINE-xx extraction, trainee row exclusion, zero-production row retention
 *  4. Row-level formula calculations (Mc Hrs, Earn Minutes, Invested Min, Effi)
 *  5. Buyer group normalization and mapping
 *  6. Red-flag detection (unknown buyers, helper HP decision, zero-production days)
 *  7. Validations and weekly buyer-group summary pivot
 */

const XLSX = require('xlsx');

// Normalizes whitespace and casing
function normalizeText(str) {
  if (str == null) return '';
  return String(str).trim().replace(/\s+/g, ' ');
}

function normalizeKey(str) {
  return normalizeText(str).toLowerCase();
}

/**
 * Parses date string in day-month-year format or Excel date numbers.
 * Supported: 14.9.2026, 14.09.2026, 14-9-2026, 14/9/2026, 14-Sep-2026, etc.
 * Returns 'YYYY-MM-DD' or null if unparseable.
 */
function parseTabDate(val) {
  if (val == null) return null;

  // If already Date instance
  if (val instanceof Date && !isNaN(val.getTime())) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // If Excel serial date number
  if (typeof val === 'number') {
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed) {
      const y = parsed.y;
      const m = String(parsed.m).padStart(2, '0');
      const d = String(parsed.d).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }

  const str = String(val).trim();
  // Match cell text like "DATE-14.9.2026", "DATE: 14/09/2026", "DATE - 14-9-2026"
  const dateMatch = str.match(/(?:DATE\s*[-:—–]?\s*)?(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})/i);
  if (dateMatch) {
    const day = parseInt(dateMatch[1], 10);
    const month = parseInt(dateMatch[2], 10);
    let year = parseInt(dateMatch[3], 10);
    if (year < 100) year += 2000;

    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // Also support month names e.g. "14-Sep-2026"
  const monthNames = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const textMonthMatch = str.match(/(\d{1,2})[\.\-\/\s]+([a-zA-Z]{3,})[\.\-\/\s]+(\d{2,4})/i);
  if (textMonthMatch) {
    const day = parseInt(textMonthMatch[1], 10);
    const mStr = textMonthMatch[2].toLowerCase().slice(0, 3);
    const month = monthNames[mStr];
    let year = parseInt(textMonthMatch[3], 10);
    if (year < 100) year += 2000;
    if (month && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  return null;
}

/**
 * Scans a single sheet's raw 2D array to find the date cell in the first ~8 rows.
 */
function extractSheetDate(rawSheetData) {
  const maxRows = Math.min(rawSheetData.length, 10);
  for (let r = 0; r < maxRows; r++) {
    const row = rawSheetData[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell != null && String(cell).toUpperCase().includes('DATE')) {
        const d = parseTabDate(cell);
        if (d) return { date: d, cellCoord: `R${r + 1}C${c + 1}`, rawText: String(cell) };
      }
    }
  }
  return null;
}

/**
 * Scans all tabs of an uploaded raw workbook and categorizes them:
 * - validTabs: [{ sheetName, date, rawSheetData }]
 * - ignoredTabs: [{ sheetName, reason }]
 * - dateDuplicates: { [date]: [sheetName1, sheetName2] }
 */
function scanWorkbookTabs(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const validTabs = [];
  const ignoredTabs = [];
  const dateMap = new Map();

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) {
      ignoredTabs.push({ sheetName, reason: 'Empty sheet' });
      continue;
    }

    const rawSheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    if (!rawSheetData || rawSheetData.length < 5) {
      ignoredTabs.push({ sheetName, reason: 'Too few rows to be a production sheet' });
      continue;
    }

    const dateInfo = extractSheetDate(rawSheetData);
    if (!dateInfo) {
      ignoredTabs.push({ sheetName, reason: 'No DATE cell found in first 10 rows' });
      continue;
    }

    const tabObj = {
      sheetName,
      date: dateInfo.date,
      rawDateText: dateInfo.rawText,
      rawSheetData,
    };

    validTabs.push(tabObj);

    if (!dateMap.has(dateInfo.date)) {
      dateMap.set(dateInfo.date, []);
    }
    dateMap.get(dateInfo.date).push(sheetName);
  }

  // Sort valid tabs by date ascending
  validTabs.sort((a, b) => a.date.localeCompare(b.date));

  const duplicateDates = {};
  for (const [d, tabs] of dateMap.entries()) {
    if (tabs.length > 1) {
      duplicateDates[d] = tabs;
    }
  }

  return { validTabs, ignoredTabs, duplicateDates };
}

/**
 * Header column aliases mapping normalized header -> field key
 */
const REQUIRED_HEADERS = {
  operators_op: [/no\s*\.?\s*of\s*op/i, /^op$/i, /^operators$/i],
  helpers_hp: [/no\s*\.?\s*of\s*hp/i, /^hp$/i, /^helpers$/i],
  reg_hours: [/^wh$/i, /^working\s*hours$/i, /^reg\s*hours$/i],
  ot_man: [/ot\s*man/i],
  ot_hr: [/ot\s*hr/i, /ot\s*hours/i],
  raw_buyer: [/^buyer$/i],
  style: [/^style$/i],
  item_no: [/^item$/i, /^item\s*no\.?$/i],
  order_qty: [/order\s*qty/i],
  smv: [/^smv$/i],
  production: [/\(?actual\s*sew\s*pro\)?/i, /^actual\s*sew\s*pro$/i, /^production$/i],
};

function matchHeaderCol(headerText) {
  if (!headerText) return null;
  const norm = normalizeText(headerText);
  for (const [field, patterns] of Object.entries(REQUIRED_HEADERS)) {
    for (const pat of patterns) {
      if (pat.test(norm)) return field;
    }
  }
  return null;
}

/**
 * Finds all unit tables in a sheet and parses their line rows.
 * Unit mapping: titleNumber (10) -> unit_code ("B10").
 */
function parseSheetLines(sheetObj, unitMap = new Map()) {
  const { sheetName, date, rawSheetData } = sheetObj;
  const rows = [];
  const changeLogs = [];
  const validations = [];

  let traineeCount = 0;
  let spareLineCount = 0;
  let cachedSheetTotalProd = null;

  let r = 0;
  while (r < rawSheetData.length) {
    const row = rawSheetData[r] || [];
    const firstCell = normalizeText(row[0]);

    // Check for Unit Table Title e.g. "Daily Production Forecast SFL#10"
    const titleMatch = firstCell.match(/Daily Production Forecast\s+SFL#(\d+)/i);
    if (titleMatch) {
      const titleNum = parseInt(titleMatch[1], 10);
      const unitCode = unitMap.get(titleNum) || `B${titleNum}`;

      // Search for the header row (next row where col A is 'Line')
      r++;
      let headerRowIdx = -1;
      while (r < rawSheetData.length) {
        const potentialHeader = normalizeText((rawSheetData[r] || [])[0]);
        if (/^Line$/i.test(potentialHeader)) {
          headerRowIdx = r;
          break;
        }
        // Stop if we hit another unit title
        if (/Daily Production Forecast\s+SFL#\d+/i.test(potentialHeader)) {
          break;
        }
        r++;
      }

      if (headerRowIdx === -1) {
        validations.push({
          date,
          sheetName,
          type: 'warning',
          message: `Could not find header row ('Line' in col A) for unit ${unitCode}`,
        });
        continue;
      }

      const headers = rawSheetData[headerRowIdx] || [];
      const colFieldMap = headers.map((h) => matchHeaderCol(h));

      // Parse data rows under this unit table
      r = headerRowIdx + 1;
      while (r < rawSheetData.length) {
        const dataRow = rawSheetData[r] || [];
        const colA = normalizeText(dataRow[0]);

        // Stop if next unit title or summary section
        if (/Daily Production Forecast\s+SFL#\d+/i.test(colA)) {
          break;
        }
        if (/^Both BG Today Total/i.test(colA) || /^Till Date/i.test(colA) || /^Grand Total/i.test(colA)) {
          break;
        }

        // Check for Trainee row: "New Join & Additional"
        if (/^New Join/i.test(colA)) {
          traineeCount++;
          r++;
          continue;
        }

        // Check for Total row
        if (/^Total:/i.test(colA) || (colA === '' && dataRow.some((c) => String(c).includes('Total')))) {
          // Attempt to capture cached sheet total production
          colFieldMap.forEach((field, idx) => {
            if (field === 'production' && dataRow[idx] != null && !isNaN(Number(dataRow[idx]))) {
              cachedSheetTotalProd = (cachedSheetTotalProd || 0) + Number(dataRow[idx]);
            }
          });
          r++;
          continue;
        }

        // Check if data row is a LINE row: /^LINE-\d+$/i
        const lineMatch = colA.match(/^LINE-(\d+)$/i);
        if (lineMatch) {
          const lineCode = `LINE-${String(lineMatch[1]).padStart(2, '0')}`;
          const obj = {
            work_date: date,
            line: lineCode,
            unit: unitCode,
            source_tab: sheetName,
            source_row_no: r + 1,
            operators_op: 0,
            helpers_hp: 0,
            reg_hours: 8,
            ot_man: 0,
            ot_hr: 0,
            raw_buyer: '',
            style: '',
            item_no: '',
            order_qty: 0,
            smv: 0,
            production: 0,
          };

          colFieldMap.forEach((field, cIdx) => {
            if (!field) return;
            const cellVal = dataRow[cIdx];
            if (cellVal == null || cellVal === '') return;

            if (field === 'raw_buyer' || field === 'style') {
              obj[field] = normalizeText(cellVal);
            } else if (field === 'item_no') {
              // Item no. must remain TEXT untouched (never zero-pad or convert numbers)
              obj.item_no = String(cellVal).trim();
            } else {
              // Numeric fields
              const n = Number(cellVal);
              obj[field] = isNaN(n) ? 0 : n;
            }
          });

          // Check for spare line: Operators is 0/empty AND Buyer is empty AND Production is 0/empty
          if ((!obj.operators_op || obj.operators_op === 0) && (!obj.raw_buyer || obj.raw_buyer === '') && (!obj.production || obj.production === 0)) {
            spareLineCount++;
            r++;
            continue;
          }

          rows.push(obj);
        }

        r++;
      }
    } else {
      r++;
    }
  }

  if (traineeCount > 0) {
    changeLogs.push({
      category: 'trainee_excluded',
      details: { date, count: traineeCount, sheetName },
    });
  }

  return { rows, changeLogs, validations, cachedSheetTotalProd, spareLineCount };
}

/**
 * Computes Mc Hrs, Earn Minutes, Invested Min, and Effi for a row.
 */
function calculateRowMetrics(row, includeHp = false) {
  const op = Number(row.operators_op) || 0;
  const hp = Number(row.helpers_hp) || 0;
  const opsFinal = includeHp ? (op + hp) : op;
  row.operators_final = opsFinal;

  const regHours = row.reg_hours != null && !isNaN(Number(row.reg_hours)) ? Number(row.reg_hours) : 8;
  const otMan = Number(row.ot_man) || 0;
  const otHr = Number(row.ot_hr) || 0;
  const smv = Number(row.smv) || 0;
  const prod = Number(row.production) || 0;

  const mcHrs = (opsFinal * regHours) + (otMan * otHr);
  const earnMinutes = smv * prod;
  const investedMin = mcHrs * 60;
  const effi = investedMin > 0 ? (earnMinutes / investedMin) : 0;

  row.mc_hrs = mcHrs;
  row.earn_minutes = earnMinutes;
  row.effi = effi;
}

/**
 * Resolves buyer groups from normalized raw buyer strings.
 * buyerMap: Map<normalized_raw_buyer, { group_name, buyer_group_id }>
 */
function resolveRowBuyer(row, buyerMap = new Map()) {
  const raw = normalizeText(row.raw_buyer);
  const normKey = normalizeKey(raw);

  if (!raw) {
    row.buyer_group = null;
    row.is_pending = true;
    row.pending_reason = 'Missing buyer on line with operators';
    return;
  }

  const mapping = buyerMap.get(normKey);
  if (mapping) {
    row.buyer_group = mapping.group_name || mapping;
    row.is_pending = false;
    row.pending_reason = null;
  } else {
    row.buyer_group = null;
    row.is_pending = true;
    row.pending_reason = `New buyer '${raw}' – needs mapping`;
  }
}

/**
 * Full Pipeline for processing a raw workbook across a chosen date range.
 *
 * @param {Buffer} buffer - Excel workbook buffer
 * @param {Object} options
 * @param {string} options.startDate - 'YYYY-MM-DD'
 * @param {string} options.endDate - 'YYYY-MM-DD'
 * @param {number|null} options.workingDays - User-provided working days (optional)
 * @param {boolean|null} options.includeHp - Decision for helpers (null = undecided)
 * @param {Map} options.buyerMap - Normalized raw buyer -> group
 * @param {Map} options.unitMap - SFL title number -> unit code (10->B10, 11->B11)
 */
function processWorkbook(buffer, { startDate, endDate, workingDays, includeHp, buyerMap = new Map(), unitMap = new Map() }) {
  const { validTabs, ignoredTabs, duplicateDates } = scanWorkbookTabs(buffer);

  // Filter valid tabs strictly within the requested date range
  const inRangeTabs = validTabs.filter((t) => t.date >= startDate && t.date <= endDate);
  const outOfRangeTabs = validTabs.filter((t) => t.date < startDate || t.date > endDate);

  const foundDates = inRangeTabs.map((t) => t.date);
  const uniqueFoundDates = Array.from(new Set(foundDates)).sort();

  // Compute missing dates in range
  const allRangeDates = [];
  const curr = new Date(startDate);
  const end = new Date(endDate);
  while (curr <= end) {
    const y = curr.getFullYear();
    const m = String(curr.getMonth() + 1).padStart(2, '0');
    const d = String(curr.getDate()).padStart(2, '0');
    allRangeDates.push(`${y}-${m}-${d}`);
    curr.setDate(curr.getDate() + 1);
  }
  const missingDates = allRangeDates.filter((d) => !uniqueFoundDates.includes(d));

  // Determine working days
  const autoWorkingDays = uniqueFoundDates.length;
  const finalWorkingDays = (workingDays && Number(workingDays) > 0) ? Number(workingDays) : autoWorkingDays;

  const allRows = [];
  const allChangeLogs = [];
  const allValidations = [];
  let totalHpCount = 0;
  let totalZeroProdRows = 0;
  let totalZeroProdOperators = 0;

  // Record ignored tabs
  for (const ign of ignoredTabs) {
    allChangeLogs.push({
      category: 'ignored_tab',
      details: { tab: ign.sheetName, reason: ign.reason },
    });
  }

  // Parse sheets in range
  for (const sheetObj of inRangeTabs) {
    const { rows, changeLogs, validations, cachedSheetTotalProd } = parseSheetLines(sheetObj, unitMap);
    allChangeLogs.push(...changeLogs);
    allValidations.push(...validations);

    // Validation: check daily production against cached total if available
    let dayTotalProd = 0;
    for (const r of rows) {
      dayTotalProd += Number(r.production) || 0;
    }

    if (cachedSheetTotalProd != null && Math.abs(dayTotalProd - cachedSheetTotalProd) > 0.01) {
      allValidations.push({
        date: sheetObj.date,
        type: 'warning',
        message: `Parsed production sum (${dayTotalProd}) does not match tab cached total (${cachedSheetTotalProd}) on ${sheetObj.date}`,
      });
    }

    // Validation: expected 35 lines per day (20 + 15)
    if (rows.length !== 35) {
      allValidations.push({
        date: sheetObj.date,
        type: 'info',
        message: `Found ${rows.length} LINE rows on ${sheetObj.date} (standard is 35 lines)`,
      });
    }

    // Process each row
    for (const row of rows) {
      if ((Number(row.helpers_hp) || 0) > 0) {
        totalHpCount++;
      }

      // Check zero-production rows
      const prod = Number(row.production) || 0;
      if (prod === 0) {
        totalZeroProdRows++;
        totalZeroProdOperators += (Number(row.operators_op) || 0);
      }

      // Resolve buyer
      resolveRowBuyer(row, buyerMap);

      // Compute metrics
      calculateRowMetrics(row, includeHp === true);

      // Row-level validations
      if (prod > 0 && (!row.style || row.smv === 0 || row.order_qty === 0)) {
        allValidations.push({
          date: row.work_date,
          line: row.line,
          unit: row.unit,
          type: 'warning',
          message: `Missing Style/SMV/OrderQty on ${row.unit} ${row.line} with production ${prod}`,
        });
      }

      if (row.effi > 1.0) {
        allValidations.push({
          date: row.work_date,
          line: row.line,
          unit: row.unit,
          type: 'warning',
          message: `Efficiency > 100% (${(row.effi * 100).toFixed(1)}%) on ${row.unit} ${row.line} (${row.work_date})`,
        });
      }

      if ((row.ot_hr > 0 && row.ot_man === 0) || (row.ot_man > 0 && row.ot_hr === 0)) {
        allValidations.push({
          date: row.work_date,
          line: row.line,
          unit: row.unit,
          type: 'warning',
          message: `Mismatched OT: OT Man = ${row.ot_man}, OT Hr = ${row.ot_hr} on ${row.unit} ${row.line} (${row.work_date})`,
        });
      }

      allRows.push(row);
    }
  }

  // Sort rows: Date ascending -> Unit (B10 first, then B11) -> Line ascending
  allRows.sort((a, b) => {
    if (a.work_date !== b.work_date) return a.work_date.localeCompare(b.work_date);
    if (a.unit !== b.unit) return a.unit.localeCompare(b.unit);
    return a.line.localeCompare(b.line);
  });

  if (totalZeroProdRows > 0) {
    allChangeLogs.push({
      category: 'zero_prod_kept',
      details: { row_count: totalZeroProdRows, operator_count: totalZeroProdOperators },
    });
  }

  if (totalHpCount > 0) {
    allChangeLogs.push({
      category: 'hp_found',
      details: { count: totalHpCount, include_hp: includeHp },
    });
  }

  // Pending issues collection
  const unknownBuyersMap = new Map();
  for (const r of allRows) {
    if (r.is_pending && r.raw_buyer) {
      const key = r.raw_buyer;
      if (!unknownBuyersMap.has(key)) {
        unknownBuyersMap.set(key, { raw_buyer: key, count: 0, dates: new Set(), lines: new Set() });
      }
      const item = unknownBuyersMap.get(key);
      item.count++;
      item.dates.add(r.work_date);
      item.lines.add(`${r.unit} ${r.line}`);
    }
  }

  const unknownBuyers = Array.from(unknownBuyersMap.values()).map((ub) => ({
    raw_buyer: ub.raw_buyer,
    count: ub.count,
    dates: Array.from(ub.dates).sort(),
    lines: Array.from(ub.lines).slice(0, 5),
  }));

  const pendingHpDecision = totalHpCount > 0 && includeHp == null;

  return {
    rows: allRows,
    foundDates: uniqueFoundDates,
    missingDates,
    autoWorkingDays,
    finalWorkingDays,
    changeLogs: allChangeLogs,
    validations: allValidations,
    unknownBuyers,
    totalHpCount,
    pendingHpDecision,
    totalZeroProdRows,
    totalZeroProdOperators,
    duplicateDates,
  };
}

/**
 * Builds the final buyer group summary pivot from merged rows.
 */
function buildFinalSummary(rows, workingDays) {
  const groupMap = new Map();

  for (const r of rows) {
    const group = r.buyer_group || 'Unassigned';
    if (!groupMap.has(group)) {
      groupMap.set(group, {
        buyer_group: group,
        sum_production: 0,
        sum_operators: 0,
      });
    }
    const item = groupMap.get(group);
    item.sum_production += Number(r.production) || 0;
    item.sum_operators += Number(r.operators_final) || 0;
  }

  // Alphabetical order
  const summary = Array.from(groupMap.values()).sort((a, b) => a.buyer_group.localeCompare(b.buyer_group));

  for (const item of summary) {
    item.avg_prod_per_day = workingDays > 0 ? (item.sum_production / workingDays) : 0;
    item.avg_optrs_per_day = workingDays > 0 ? (item.sum_operators / workingDays) : 0;
  }

  const grandTotal = {
    buyer_group: 'Total',
    sum_production: summary.reduce((s, g) => s + g.sum_production, 0),
    sum_operators: summary.reduce((s, g) => s + g.sum_operators, 0),
    avg_prod_per_day: 0,
    avg_optrs_per_day: 0,
  };
  grandTotal.avg_prod_per_day = workingDays > 0 ? (grandTotal.sum_production / workingDays) : 0;
  grandTotal.avg_optrs_per_day = workingDays > 0 ? (grandTotal.sum_operators / workingDays) : 0;

  return { summary, grandTotal };
}

module.exports = {
  normalizeText,
  normalizeKey,
  parseTabDate,
  extractSheetDate,
  scanWorkbookTabs,
  parseSheetLines,
  calculateRowMetrics,
  resolveRowBuyer,
  processWorkbook,
  buildFinalSummary,
};
