const XLSX = require('xlsx');

const OUTPUT_HEADER = ['O.CODE', '::', '', '', '', 'O.PROD', 'O.UDITEM', 'O.CUST', 'O^DD:1', 'O^DQ:1', 'O.CONTRACT_QTY', 'O.UDSEASON', 'O.UDPCD', 'O.STATUS', 'O.effy', 'O.Descrip', 'O.UDBuy details', 'O.Sprice', 'O.Udconversion', 'END'];

function normalizedHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findColumn(row, name) {
  const wanted = normalizedHeader(name);
  const key = Object.keys(row).find((candidate) => normalizedHeader(candidate) === wanted);
  return key === undefined ? '' : row[key];
}

function text(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function quantity(value) {
  if (value === undefined || value === null || text(value) === '') return null;
  const parsed = Number(text(value).replace(/,/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function excelDateToParts(value) {
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value);
    return date ? { year: date.y, month: date.m, day: date.d } : null;
  }
  const valueText = text(value);
  if (!valueText || valueText.toLowerCase() === 'same as above') return null;
  let match = valueText.match(/^(\d{1,2})[-\\/](\d{1,2})[-\\/](\d{4})$/);
  if (match) return { year: Number(match[3]), month: Number(match[2]), day: Number(match[1]) };
  match = valueText.match(/^(\d{4})[-\\/](\d{1,2})[-\\/](\d{1,2})$/);
  if (match) return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };
  match = valueText.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (match) return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };
  return null;
}

function formatDate(value, compact = false) {
  const date = excelDateToParts(value);
  if (!date) return '';
  const day = String(date.day).padStart(2, '0');
  const month = String(date.month).padStart(2, '0');
  const year = date.year;
  return compact ? `${year}${month}${day}` : `${day}-${month}-${year}`;
}

function dateWithFallback(value, previousValue, fallback = '') {
  const direct = formatDate(value);
  if (direct) return direct;
  if (text(value).toLowerCase() === 'same as above' && previousValue) return previousValue;
  return fallback;
}

function categoryFor(description) {
  const tokens = text(description).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
  if (['JACKET', 'JKT', 'BLAZER'].some((keyword) => tokens.some((token) => token.includes(keyword)))) return 'JACKET';
  const keywords = ['SHIRT', 'TEE', 'POLO', 'TOP', 'TROUSER', 'PANT', 'JEAN', 'SHORT', 'SKIRT', 'LEGGING', 'BOTTOM', 'DRESS', 'SHRT', 'PNT', 'PT', 'SH', 'SK', 'PA'];
  return keywords.some((keyword) => tokens.some((token) => token.includes(keyword))) ? 'SHIRT & BOTTOM' : '';
}

function efficiency(category, qty) {
  if (qty === null || !category) return '';
  const slabs = category === 'JACKET'
    ? [[1000, 10], [1500, 23], [3000, 30], [6500, 35], [10000, 42], [20000, 50], [30000, 58], [50000, 63], [Infinity, 65]]
    : [[1000, 10], [1500, 24], [3000, 33], [6500, 40], [10000, 44], [20000, 51], [30000, 58], [50000, 63], [100000, 65], [Infinity, 70]];
  return slabs.find(([upperBound]) => qty <= upperBound)[1];
}

function csvCell(value) {
  const cell = value === null || value === undefined ? '' : String(value);
  return /[",\r\n]/.test(cell) ? `"${cell.replace(/"/g, '""')}"` : cell;
}

function orderSeason(value) {
  const season = text(value).replace(/['’]27$/i, '');
  return /^(FA|HO|SP|SU)$/i.test(season) ? `${season}'27` : season;
}

function generateOrderFile(buffer, weekCode = 'W36') {
  // Keep Excel date cells as serial numbers so no local timezone is applied.
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: true });
  if (!workbook.SheetNames.includes('Summary sheet')) throw new Error('The uploaded workbook must contain a sheet named "Summary sheet".');
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets['Summary sheet'], { defval: '' });
  let previousPcd = '';
  let previousRevisedPcd = '';
  let previousExFactory = '';
  let previousBuyDate = '';
  const codeOccurrences = new Map();
  const output = [];
  for (const row of rows) {
    const qty = quantity(findColumn(row, 'Qty'));
    if (qty === null) continue;
    const style = text(findColumn(row, 'Style'));
    const season = text(findColumn(row, 'SEASON'));
    const pcd = dateWithFallback(findColumn(row, 'PCD'), previousPcd);
    const revisedPcd = dateWithFallback(findColumn(row, 'Revised PCD'), previousRevisedPcd, pcd);
    const exFactory = dateWithFallback(findColumn(row, 'EX FACTORY'), previousExFactory, revisedPcd || pcd);
    const buyDate = dateWithFallback(findColumn(row, 'Buy date'), previousBuyDate);
    previousPcd = pcd || previousPcd;
    previousRevisedPcd = revisedPcd || previousRevisedPcd;
    previousExFactory = exFactory || previousExFactory;
    previousBuyDate = buyDate || previousBuyDate;
    const codeSeason = orderSeason(season);
    const codeDate = exFactory ? exFactory.split('-').reverse().join('') : '';
    const baseCode = `${style}::${codeSeason}::${codeDate}::${weekCode}`;
    const occurrence = codeOccurrences.get(baseCode) || 0;
    codeOccurrences.set(baseCode, occurrence + 1);
    const orderCode = occurrence === 0
      ? baseCode
      : `${style}::${codeSeason}::${codeDate}_${occurrence}::${weekCode}`;
    const description = findColumn(row, 'Description');
    output.push([
      orderCode, '::', codeDate, weekCode, '', style, style, 'NIKE',
      exFactory, Math.round(qty * 1.03), qty, codeSeason, revisedPcd, 'P', efficiency(categoryFor(description), qty), description,
      buyDate, '', 75, 'END',
    ].map(csvCell).join(','));
  }
  return `${OUTPUT_HEADER.map(csvCell).join(',')}\r\n${output.join('\r\n')}${output.length ? '\r\n' : ''}`;
}

module.exports = { generateOrderFile, categoryFor, efficiency, formatDate };