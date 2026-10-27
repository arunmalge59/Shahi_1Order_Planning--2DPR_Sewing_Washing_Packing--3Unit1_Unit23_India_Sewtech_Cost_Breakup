const XLSX = require('xlsx');

// Internal field keys — used to read/write each record. Kept unique (even for the
// blank "Unnamed" columns) so they can be used as plain JS object keys.
const OUTPUT_FIELD_KEYS = [
  'O.CODE', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5',
  'O.PROD', 'O.UDITEM', 'O.CUST', 'O^DD:1', 'O^DQ:1', 'O.CONTRACT_QTY',
  'O.UDSEASON', 'O.UDPCD', 'O.STATUS', 'O.effy', 'O.Descrip', 'O.UDBuy details',
  'O.Sprice', 'O.Udconversion', 'END',
];

// What actually gets written into row 1 of the CSV. The "Unnamed" columns must be
// blank in the header row (matching Correct_sheet_Orders-H&M.csv) — only their
// data rows below are populated.
const OUTPUT_DISPLAY_HEADERS = OUTPUT_FIELD_KEYS.map((key) => (key.startsWith('Unnamed:') ? '' : key));

const CUSTOMER = process.env.HM_CUSTOMER || 'H&M- & DENIM';

// Columns the FR workbook must contain. "Order description" is read too but is
// NOT used for O.Descrip (confirmed against Correct_sheet_Orders-H&M.csv —
// O.Descrip comes from "FR Product", not "Order description").
const FR_REQUIRED_COLUMNS = ['Season', 'Style Proj.', 'FR Style', 'FR Product'];

function normalizedHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeText(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

const text = normalizeText;

// Used only for MATCHING (Season + Style lookup key), never for display. Collapses
// internal whitespace and lowercases so trivial spacing/case differences between
// the planning sheet and the FR sheet don't cause a false "no match".
function matchKey(...parts) {
  return parts.map((part) => text(part).replace(/\s+/g, ' ').toLowerCase()).join('||');
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

function normalizeExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value);
    return date ? { year: date.y, month: date.m, day: date.d } : null;
  }
  const valueText = text(value);
  let match = valueText.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (match) return { year: +match[3], month: +match[2], day: +match[1] };
  match = valueText.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (match) return { year: +match[1], month: +match[2], day: +match[3] };
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

// Generic "find the sheet + header row that contains these required columns" scan,
// shared by the H&M planning workbook and the FR style workbook.
function findRowsWithHeaders(workbook, requiredColumns, carryForwardHeaderKeys) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    const headerIndex = matrix.findIndex((row) => requiredColumns.every((column) => row.some((key) => normalizedHeader(key) === normalizedHeader(column))));
    if (headerIndex >= 0) {
      const headers = matrix[headerIndex].map((header) => text(header));
      const carryForwardColumns = new Set(carryForwardHeaderKeys || []);
      const previousValues = {};
      const rows = matrix.slice(headerIndex + 1).map((values, index) => {
        const row = { sourceRowNumber: headerIndex + index + 2 };
        headers.forEach((header, columnIndex) => {
          if (!header) return;
          const value = values[columnIndex] ?? '';
          const headerKey = normalizedHeader(header);
          if (value === '' && carryForwardColumns.has(headerKey) && previousValues[headerKey] !== undefined) {
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
  return null;
}

function findSourceRows(workbook) {
  const carryForwardColumns = ['status', 'season', 'style', 'style2', 'item', 'unit', 'productdescription'];
  const found = findRowsWithHeaders(workbook, ['Status', 'Qty', 'Style', 'Proj Ex-fty', 'Proj PCD'], carryForwardColumns);
  if (!found) throw new Error('No worksheet with the required H&M columns was found.');
  return found;
}

// Builds a Map of "Season + Style" -> { frStyle, frProduct, orderDescription } from
// the FR style workbook (e.g. H_M_Denim_FR_style.xlsx). Returns { lookup, sheetName,
// rowCount } so the caller can report how many FR rows were loaded.
function loadFrStyleLookup(frBuffer) {
  const workbook = XLSX.read(frBuffer, { type: 'buffer', cellDates: false, raw: true });
  const found = findRowsWithHeaders(workbook, FR_REQUIRED_COLUMNS, ['season']);
  if (!found) {
    throw new Error('The FR file is missing one of the required columns: Season, Style Proj., FR Style, FR Product.');
  }
  const lookup = new Map();
  for (const row of found.rows) {
    const season = getColumn(row, 'Season');
    const styleProj = getColumn(row, 'Style Proj.');
    const frStyle = text(getColumn(row, 'FR Style'));
    if (!season || !styleProj || !frStyle) continue;
    lookup.set(matchKey(season, styleProj), {
      frStyle,
      frProduct: text(getColumn(row, 'FR Product')),
      orderDescription: text(getColumn(row, 'Order description')),
    });
  }
  return { lookup, sheetName: found.sheetName, rowCount: found.rows.length };
}

function productIdentity(row) {
  return text(getColumn(row, 'Style')) || text(getColumn(row, 'Style2')) || text(getColumn(row, 'Item'));
}

function normalizeUnit(value) {
  const unit = text(value).toUpperCase();
  const match = unit.match(/^UNIT\s*(\d+)$/);
  return match ? `U${match[1]}` : unit;
}

function buildHmOrderCode(row, weekCode, productOverride) {
  const product = productOverride || productIdentity(row);
  const unit = normalizeUnit(getColumn(row, 'Unit'));
  const exFactory = getColumn(row, 'Proj Ex-fty');
  return `${product}::${weekCode}::${unit}::Projection::${formatCodeDate(exFactory)}`;
}

// Appends _1, _2, _3... to any O.CODE that repeats. The FIRST occurrence of a
// given code is left untouched; only the 2nd, 3rd, ... occurrences get a suffix.
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

// frBuffer is OPTIONAL. When omitted, behavior is identical to before: O.PROD /
// O.UDITEM come from the raw Style text and O.Descrip comes from the source
// "Product Description" column, and no row is skipped for lacking an FR match.
//
// When frBuffer IS provided, every Projection row is looked up by Season + Style
// against the FR sheet's Season + Style Proj.:
//   - match found    -> O.PROD / O.UDITEM = FR Style, O.Descrip = FR Product
//   - no match found -> row is skipped (not written to the CSV) and reported in
//                        validationErrors / frMissingStyles, since a style that
//                        isn't in the FR sheet yet has no correct FR name to use.
//                        Add the style to the FR sheet and re-run to include it.
function generateHmOrderFile(buffer, weekCode, frBuffer) {
  const normalizedWeekCode = text(weekCode).toUpperCase();
  if (!/^W\d{1,2}$/.test(normalizedWeekCode)) throw new Error('Week Code must be in the format W33.');
  // slice(1) preserves any leading zero, so a UI value of "W01" becomes "WK-01".
  const erpWeekCode = `WK-${normalizedWeekCode.slice(1)}`;

  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: true });
  const { sheetName, rows } = findSourceRows(workbook);
  const confirmedRows = rows.filter((row) => text(getColumn(row, 'Status')).toLowerCase() === 'confirmed').length;
  const projectionRows = rows.filter((row) => text(getColumn(row, 'Status')).toLowerCase() === 'projection');
  if (!projectionRows.length) throw new Error('No Projection rows found in the uploaded workbook.');

  let frInfo = null;
  if (frBuffer) {
    frInfo = loadFrStyleLookup(frBuffer);
  }

  const validationErrors = [];
  const frMissingStyles = [];
  const records = [];
  let frMatchedRows = 0;

  for (const row of projectionRows) {
    const qtyValue = getColumn(row, 'Qty');
    const exFactoryValue = getColumn(row, 'Proj Ex-fty');
    const pcdValue = getColumn(row, 'Proj PCD');
    const styleValue = productIdentity(row);
    const qty = normalizeQuantity(qtyValue);
    const exFactory = formatOrderDate(getColumn(row, 'Proj Ex-fty'));
    const pcd = formatOrderDate(getColumn(row, 'Proj PCD'));
    const product = styleValue;
    const reasons = [];
    if (qty === null) reasons.push('Invalid Qty');
    if (!product) reasons.push('Missing Style');
    if (!exFactory) reasons.push('Missing or invalid Proj Ex-fty');
    if (!pcd) reasons.push('Missing or invalid Proj PCD');
    if (reasons.length) {
      const validationError = {
        sourceRowNumber: row.sourceRowNumber,
        status: text(getColumn(row, 'Status')),
        style: text(styleValue),
        qty: text(qtyValue),
        projExFty: text(exFactoryValue),
        projPcd: text(pcdValue),
        reason: reasons.join('; '),
        type: 'data',
      };
      console.warn('Skipped invalid H&M Projection row:', validationError);
      validationErrors.push(validationError);
      continue;
    }

    const season = text(getColumn(row, 'Season'));
    const unit = normalizeUnit(getColumn(row, 'Unit'));

    // Resolve the FR style/description, if an FR file was supplied.
    let finalProduct = product;
    let finalDescrip = text(getColumn(row, 'Product Description'));
    if (frInfo) {
      const frMatch = frInfo.lookup.get(matchKey(season, product));
      if (!frMatch) {
        const missing = {
          sourceRowNumber: row.sourceRowNumber,
          status: 'Projection',
          style: product,
          qty: text(qtyValue),
          projExFty: text(exFactoryValue),
          projPcd: text(pcdValue),
          reason: 'No matching FR style found for this Season + Style — add it to the FR sheet and regenerate.',
          type: 'fr-missing',
        };
        validationErrors.push(missing);
        frMissingStyles.push({ season, style: product });
        continue;
      }
      finalProduct = frMatch.frStyle;
      finalDescrip = frMatch.frProduct;
      frMatchedRows += 1;
    }

    // O.CONTRACT_QTY is the exact order quantity from the source sheet (just rounded
    // to a whole unit). O^DQ:1 is 3% MORE than the contract qty (buffer/cushion qty).
    const contractQty = Math.round(qty);
    const orderQty = Math.round(contractQty * 1.03);

    records.push({
      'O.CODE': buildHmOrderCode(row, erpWeekCode, finalProduct),
      'Unnamed: 1': '::', 'Unnamed: 2': erpWeekCode, 'Unnamed: 3': 'Projection',
      'Unnamed: 4': unit, 'Unnamed: 5': formatCodeDate(getColumn(row, 'Proj Ex-fty')),
      'O.PROD': finalProduct, 'O.UDITEM': finalProduct, 'O.CUST': CUSTOMER, 'O^DD:1': exFactory,
      'O^DQ:1': orderQty, 'O.CONTRACT_QTY': contractQty, 'O.UDSEASON': season,
      'O.UDPCD': pcd, 'O.STATUS': 'P', 'O.effy': '',
      'O.Descrip': finalDescrip, 'O.UDBuy details': '',
      'O.Sprice': '', 'O.Udconversion': 75, END: 'END',
    });
  }

  dedupeOrderCodes(records);

  const csv = [
    OUTPUT_DISPLAY_HEADERS.map(csvCell).join(','),
    ...records.map((record) => OUTPUT_FIELD_KEYS.map((key) => csvCell(record[key])).join(',')),
  ].join('\r\n') + '\r\n';

  return {
    csv,
    sheetName,
    summary: {
      totalRows: rows.length,
      projectionRows: projectionRows.length,
      validProjectionRows: records.length,
      confirmedRows,
      generatedOrders: records.length,
      skippedRows: validationErrors.length,
      frProvided: Boolean(frInfo),
      frRowCount: frInfo ? frInfo.rowCount : 0,
      frMatchedRows,
      frMissingRows: frMissingStyles.length,
    },
    preview: records.slice(0, 100),
    validationErrors,
    frMissingStyles,
  };
}

module.exports = {
  generateHmOrderFile, loadFrStyleLookup, normalizeText, normalizeQuantity, normalizeExcelDate,
  formatOrderDate, formatCodeDate, buildHmOrderCode, OUTPUT_FIELD_KEYS, OUTPUT_DISPLAY_HEADERS,
};