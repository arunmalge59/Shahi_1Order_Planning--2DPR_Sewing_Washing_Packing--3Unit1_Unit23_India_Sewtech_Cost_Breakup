const { parse } = require('csv-parse/sync');
const XLSX = require('xlsx');

function cleanHeader(header) {
  return header ? header.trim() : '';
}

function toNumber(value) {
  if (value === undefined || value === null) return 0;
  const cleaned = String(value).replace(/,/g, '').trim();
  if (cleaned === '' || cleaned === '-') return 0;
  const n = Number(cleaned);
  return Number.isNaN(n) ? 0 : n;
}

function cleanDate(value) {
  if (!value) return '';
  const str = String(value).trim();
  return str === '-' ? '' : str;
}

function normalizeSeason(value) {
  const season = String(value || '').trim().toUpperCase().replace(/[`’\\]/g, "'");
  const canonical = season.match(/^(FA|HO|SP|SU)\s*'?\s*(\d{2})$/);
  if (canonical) return `${canonical[1]}'${canonical[2]}`;

  const yearPrefix = season.match(/^(\d{4})\s*(FA|HO|SP|SU)$/);
  if (yearPrefix) return `${yearPrefix[2]}'${yearPrefix[1].slice(-2)}`;

  const shortCode = season.match(/^(FA|HO|SP|SU)$/);
  if (shortCode) return `${shortCode[1]}'27`;

  return season;
}

function monthFromDate(value) {
  if (!value) return '';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  return parsed.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' });
}

function findValue(row, possibleKeys) {
  for (const key of possibleKeys) {
    if (row[key] !== undefined && row[key] !== null) return row[key];
    const lowerKey = key.toLowerCase();
    for (const k of Object.keys(row)) {
      if (k.trim().toLowerCase() === lowerKey) return row[k];
    }
  }
  return '';
}

function parseSnapshotCsv(csvText) {
  const records = parse(csvText, {
    columns: (headers) => headers.map(cleanHeader),
    skip_empty_lines: true,
    relax_column_count: true,
    trim: true,
  });

  return records
    .filter((row) => {
      const style = findValue(row, ['Style', 'style']);
      return style && String(style).trim() !== '';
    })
    .map((row) => ({
      itemNumber: String(findValue(row, ['Item number', 'Item Number', 'item_number']) || '').trim(),
      status: String(findValue(row, ['Status', 'status']) || '').trim(),
      style: String(findValue(row, ['Style', 'style']) || '').trim(),
      description: String(findValue(row, ['Description', 'description']) || '').trim(),
      qty: toNumber(findValue(row, ['Qty', 'qty', 'QTY'])),
      pcd: cleanDate(findValue(row, ['PCD', 'pcd'])),
      revisedPcd: cleanDate(findValue(row, ['Revised PCD', 'revised_pcd'])),
      exFactory: cleanDate(findValue(row, ['EX FACTORY', 'Ex Factory', 'ex_factory', 'Ex-Factory'])),
      buyDate: cleanDate(findValue(row, ['Buy date', 'Buy Date', 'buy_date'])),
      remarks: String(findValue(row, ['Remarks', 'remarks']) || '').trim(),
      season: normalizeSeason(findValue(row, ['SEASON', 'Season', 'season'])),
      month: monthFromDate(findValue(row, ['PCD', 'pcd'])) || monthFromDate(findValue(row, ['EX FACTORY', 'Ex Factory', 'ex_factory', 'Ex-Factory'])) || monthFromDate(findValue(row, ['Buy date', 'Buy Date', 'buy_date'])),
    }));
}

function parseSnapshotFile(buffer, originalName = '') {
  if (/\.xlsx?$/i.test(originalName)) {
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    return parseSnapshotCsv(XLSX.utils.sheet_to_csv(firstSheet));
  }
  return parseSnapshotCsv(buffer.toString('utf-8'));
}

module.exports = { parseSnapshotCsv, parseSnapshotFile, toNumber, cleanDate, normalizeSeason };

