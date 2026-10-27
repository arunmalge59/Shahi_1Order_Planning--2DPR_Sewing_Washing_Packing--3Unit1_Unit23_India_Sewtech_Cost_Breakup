const XLSX = require('xlsx');

// Internal field keys matching the target Adidas order-file schema (21 columns)
const OUTPUT_FIELD_KEYS = [
  'O.CODE', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5',
  'O.PROD', 'O.UDITEM', 'O.CUST', 'O^DD:1', 'O^DQ:1', 'O.CONTRACT_QTY',
  'O.UDSEASON', 'O.UDPCD', 'O.STATUS', 'O.effy', 'O.Descrip', 'O.UDBuy details',
  'O.Sprice', 'O.Udconversion', 'END',
];

// CSV display headers — Unnamed columns remain empty header text
const OUTPUT_DISPLAY_HEADERS = OUTPUT_FIELD_KEYS.map((key) => (key.startsWith('Unnamed:') ? '' : key));

const ADIDAS_CUSTOMER = process.env.ADIDAS_CUSTOMER || 'ADIDAS-Originals';

function normalizedHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function text(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

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

const MONTH_MAP = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
};

function normalizeExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value);
    return date ? { year: date.y, month: date.m, day: date.d } : null;
  }
  const valueText = text(value);
  let match = valueText.match(/^(\d{1,2})[-\/]([a-zA-Z]{3}|\d{1,2})[-\/](\d{4})$/);
  if (match) {
    const day = +match[1];
    const rawMonth = match[2].toLowerCase();
    const month = MONTH_MAP[rawMonth] || +rawMonth;
    const year = +match[3];
    if (day && month && year) return { year, month, day };
  }
  match = valueText.match(/^(\d{4})[-\/]([a-zA-Z]{3}|\d{1,2})[-\/](\d{1,2})$/);
  if (match) {
    const year = +match[1];
    const rawMonth = match[2].toLowerCase();
    const month = MONTH_MAP[rawMonth] || +rawMonth;
    const day = +match[3];
    if (day && month && year) return { year, month, day };
  }
  match = valueText.match(/^(\d{4})(\d{2})(\d{2})$/);
  return match ? { year: +match[1], month: +match[2], day: +match[3] } : null;
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

// Scans sheets to locate the header row containing key Order Chart columns
function findSourceRows(workbook) {
  const requiredCandidates = ['Style Number', 'Style', 'Qty', 'Quantity', 'Status', 'PCD', 'X-fty Date', 'Xfty Date', 'Ex-fty Date'];
  
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    
    // Find header row that matches at least Style + Qty or Status
    const headerIndex = matrix.findIndex((row) => {
      const normalizedRow = row.map(normalizedHeader);
      const hasStyle = normalizedRow.includes('stylenumber') || normalizedRow.includes('style');
      const hasQty = normalizedRow.includes('qty') || normalizedRow.includes('quantity');
      return hasStyle && hasQty;
    });

    if (headerIndex >= 0) {
      const headers = matrix[headerIndex].map((header) => text(header));
      const previousValues = {};
      const carryForwardHeaderKeys = new Set(['season', 'status', 'fty', 'factory', 'unit', 'cat', 'category', 'stylenumber', 'style']);

      const rows = matrix.slice(headerIndex + 1).map((values, index) => {
        const row = { sourceRowNumber: headerIndex + index + 2 };
        headers.forEach((header, columnIndex) => {
          if (!header) return;
          const value = values[columnIndex] ?? '';
          const headerKey = normalizedHeader(header);
          if (value === '' && carryForwardHeaderKeys.has(headerKey) && previousValues[headerKey] !== undefined) {
            row[header] = previousValues[headerKey];
          } else {
            row[header] = value;
          }
          if (row[header] !== '') previousValues[headerKey] = row[header];
        });
        return row;
      }).filter((row) => Object.keys(row).some((key) => key !== 'sourceRowNumber' && text(row[key])));

      return { sheetName, rows };
    }
  }

  // Fallback: try sheet_to_json directly if header scan fails
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
  return { sheetName, rows };
}

function extractStyleNumber(row) {
  return text(getColumn(row, 'Style Number')) || text(getColumn(row, 'Style')) || text(getColumn(row, 'Item Number')) || text(getColumn(row, 'Item'));
}

function extractFactory(row) {
  const fty = text(getColumn(row, 'Fty')) || text(getColumn(row, 'Factory')) || text(getColumn(row, 'Unit'));
  if (!fty) return 'U1';
  const match = fty.match(/^UNIT\s*(\d+)$/i);
  return match ? `U${match[1]}` : fty;
}

function extractDate(row) {
  return getColumn(row, 'X-fty Date') || getColumn(row, 'Xfty Date') || getColumn(row, 'Poss X-fty') || getColumn(row, 'Extended dates') || getColumn(row, 'Ex-fty Date');
}

function extractPcd(row) {
  return getColumn(row, 'PCD') || getColumn(row, 'Plan PCD');
}

function extractCategory(row) {
  return text(getColumn(row, 'Cat')) || text(getColumn(row, 'Category')) || text(getColumn(row, 'Item Number'));
}

function formatSeasonForProd(seasonText) {
  const season = text(seasonText);
  if (!season) return "SS'27";
  if (season.includes("'")) return season;
  const match = season.match(/^([A-Za-z]+)(\d{2})$/);
  if (match) {
    return `${match[1]}'${match[2]}`;
  }
  return season;
}

function buildProdStyle(style, seasonText) {
  const formattedSeason = formatSeasonForProd(seasonText);
  if (style.endsWith(`-${formattedSeason}`) || style.endsWith(`-${seasonText}`)) {
    return style;
  }
  return `${style}-${formattedSeason}`;
}

function buildAdidasOrderCode(prodStyle, weekCode, fty, status, dateValue) {
  const codeDate = formatCodeDate(dateValue);
  return `${prodStyle}::${weekCode}::${fty}::${status}::${codeDate}`;
}


function dedupeOrderCodes(records) {
  const seenCounts = new Map();
  for (const record of records) {
    const baseCode = record['O.CODE'];
    const occurrence = (seenCounts.get(baseCode) || 0) + 1;
    seenCounts.set(baseCode, occurrence);
    if (occurrence > 1) {
      record['O.CODE'] = `${baseCode}_${occurrence - 1}`;
    }
  }
  return records;
}

function generateAdidasOrderFile(buffer, weekCode = 'WK-33 Rev1') {
  const userWeekCode = text(weekCode) || 'WK-33 Rev1';
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: true });
  const { sheetName, rows } = findSourceRows(workbook);

  const validationErrors = [];
  const validRows = [];

  for (const row of rows) {
    const style = extractStyleNumber(row);
    const rawQty = getColumn(row, 'Qty') || getColumn(row, 'Quantity');
    const qty = normalizeQuantity(rawQty);

    if (!style || qty === null || qty <= 0) {
      if (rawQty || style) {
        validationErrors.push({
          sourceRowNumber: row.sourceRowNumber || 0,
          style,
          qty: text(rawQty),
          reason: !style ? 'Missing Style Number' : 'Invalid/Zero Quantity',
        });
      }
      continue;
    }

    const season = text(getColumn(row, 'Season')) || 'SS27';
    const status = text(getColumn(row, 'Status')) || 'Confirmed';
    const fty = extractFactory(row);
    const xftyDate = extractDate(row);
    const pcdDate = extractPcd(row);
    const category = extractCategory(row);

    validRows.push({
      sourceRowNumber: row.sourceRowNumber,
      style,
      season,
      status,
      fty,
      xftyDate,
      pcdDate,
      category,
      qty,
    });
  }

  if (validRows.length === 0) {
    throw new Error('No valid order rows found in the uploaded workbook.');
  }

  // Group valid rows by unique order attributes
  const groups = new Map();

  for (const row of validRows) {
    const prodStyle = buildProdStyle(row.style, row.season);
    const exFactoryStr = formatOrderDate(row.xftyDate);
    const pcdStr = formatOrderDate(row.pcdDate);
    
    // Grouping key: style + season + fty + status + xfty + pcd + category
    const groupKey = `${prodStyle}||${row.season}||${row.fty}||${row.status}||${exFactoryStr}||${pcdStr}||${row.category.toLowerCase()}`;

    if (!groups.has(groupKey)) {
      groups.set(groupKey, {
        prodStyle,
        rawStyle: row.style,
        season: row.season,
        fty: row.fty,
        status: row.status,
        xftyDate: row.xftyDate,
        pcdDate: row.pcdDate,
        category: row.category,
        totalQty: 0,
        sourceRowCount: 0,
      });
    }

    const group = groups.get(groupKey);
    group.totalQty += row.qty;
    group.sourceRowCount += 1;
  }

  const records = [];

  for (const group of groups.values()) {
    const contractQty = Math.round(group.totalQty);
    // Order quantity includes 3% cushion allowance (O^DQ:1 = Math.round(contractQty * 1.03))
    const orderQty = Math.round(contractQty * 1.03);

    const exFactoryFormatted = formatOrderDate(group.xftyDate);
    const pcdFormatted = formatOrderDate(group.pcdDate);
    const codeDate = formatCodeDate(group.xftyDate);

    const orderCode = buildAdidasOrderCode(group.prodStyle, userWeekCode, group.fty, group.status, group.xftyDate);

    records.push({
      'O.CODE': orderCode,
      'Unnamed: 1': '::',
      'Unnamed: 2': userWeekCode,
      'Unnamed: 3': group.fty,
      'Unnamed: 4': group.status,
      'Unnamed: 5': codeDate,
      'O.PROD': group.prodStyle,
      'O.UDITEM': group.prodStyle,
      'O.CUST': ADIDAS_CUSTOMER,
      'O^DD:1': exFactoryFormatted,
      'O^DQ:1': orderQty,
      'O.CONTRACT_QTY': contractQty,
      'O.UDSEASON': group.season,
      'O.UDPCD': pcdFormatted,
      'O.STATUS': group.status === 'Confirmed' ? 'C' : (group.status === 'Projection' ? 'P' : group.status),
      'O.effy': '',
      'O.Descrip': group.category ? group.category.toUpperCase() : '',
      'O.UDBuy details': '',
      'O.Sprice': '',
      'O.Udconversion': 75,
      END: 'END',
    });
  }

  dedupeOrderCodes(records);

  const csv = [
    OUTPUT_DISPLAY_HEADERS.map(csvCell).join(','),
    ...records.map((record) => OUTPUT_FIELD_KEYS.map((key) => csvCell(record[key])).join(',')),
  ].join('\r\n') + '\r\n';

  const totalContractQty = records.reduce((sum, r) => sum + Number(r['O.CONTRACT_QTY']), 0);
  const totalOrderQty = records.reduce((sum, r) => sum + Number(r['O^DQ:1']), 0);

  return {
    csv,
    sheetName,
    summary: {
      totalRows: rows.length,
      validRows: validRows.length,
      generatedOrders: records.length,
      skippedRows: validationErrors.length,
      totalContractQty,
      totalOrderQty,
    },
    preview: records.slice(0, 100),
    validationErrors,
  };
}

module.exports = {
  generateAdidasOrderFile,
  buildProdStyle,
  buildAdidasOrderCode,
  formatSeasonForProd,
  OUTPUT_FIELD_KEYS,
  OUTPUT_DISPLAY_HEADERS,
};
