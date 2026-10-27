/**
 * Sewtech Excel I/O Service
 *
 * Generates and parses Excel files for the Sewtech workflow:
 *  1. Merged Sheet export (sheet name 'Sewtech') with 18 exact columns, red highlight flags & notes
 *  2. Final Summary Sheet export (starts at D7) with dynamic buyer groups, SUM formulas & working days note
 *  3. Modified Merged Sheet parser for re-upload workflow
 */

const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

// Helper for ordinal day formatting e.g. "14th Sep to 20th Sep"
function formatOrdinalDateRange(startDateStr, endDateStr, fallbackLabel) {
  if (!startDateStr || !endDateStr) return fallbackLabel || 'Production';

  try {
    const s = new Date(startDateStr);
    const e = new Date(endDateStr);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    const getOrdinal = (n) => {
      const s = ['th', 'st', 'nd', 'rd'];
      const v = n % 100;
      return n + (s[(v - 20) % 10] || s[v] || s[0]);
    };

    const sDay = getOrdinal(s.getDate());
    const sMon = months[s.getMonth()];
    const eDay = getOrdinal(e.getDate());
    const eMon = months[e.getMonth()];

    if (sMon === eMon) {
      return `${sDay} ${sMon} to ${eDay} ${eMon}`;
    }
    return `${sDay} ${sMon} to ${eDay} ${eMon}`;
  } catch (err) {
    return fallbackLabel || 'Production';
  }
}

const MERGED_COLUMNS = [
  { header: 'Date', key: 'work_date', width: 14 },
  { header: 'Line', key: 'line', width: 12 },
  { header: 'Unit', key: 'unit', width: 10 },
  { header: 'Operators', key: 'operators_final', width: 12, numFmt: '#,##0' },
  { header: 'ReG Hours', key: 'reg_hours', width: 12, numFmt: '#0.0' },
  { header: 'OT MAN', key: 'ot_man', width: 12, numFmt: '#,##0' },
  { header: 'OT HR', key: 'ot_hr', width: 12, numFmt: '#0.0' },
  { header: 'Raw Buyer', key: 'raw_buyer', width: 18 },
  { header: 'Buyer', key: 'buyer_group', width: 18 },
  { header: 'Style', key: 'style', width: 18 },
  { header: 'Item no.', key: 'item_no', width: 14 },
  { header: 'Order Qty', key: 'order_qty', width: 14, numFmt: '#,##0' },
  { header: 'SMV', key: 'smv', width: 12, numFmt: '#0.00' },
  { header: 'Earn Minutes', key: 'earn_minutes', width: 16, numFmt: '#,##0.00' },
  { header: 'Production', key: 'production', width: 14, numFmt: '#,##0' },
  { header: 'EFFI.', key: 'effi', width: 12, numFmt: '0.00%' },
  { header: 'Mc Hrs', key: 'mc_hrs', width: 14, numFmt: '#,##0.00' },
  { header: 'REMARKS', key: 'remarks', width: 18 },
];

/**
 * Writes the flat Merged Sheet (sheet name 'Sewtech') with red fills on pending rows.
 */
async function writeMergedSheet(rows, { weekLabel } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sewtech', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  // Add header row
  const headerRow = ws.addRow(MERGED_COLUMNS.map((c) => c.header));
  headerRow.font = { bold: true, color: { argb: 'FF000000' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

  // Set column widths
  MERGED_COLUMNS.forEach((col, idx) => {
    ws.getColumn(idx + 1).width = col.width;
  });

  // Enable AutoFilter across all columns
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: MERGED_COLUMNS.length },
  };

  const redFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFC7CE' },
  };
  const redFont = {
    color: { argb: 'FF9C0006' },
    bold: true,
  };

  for (const row of rows) {
    const rawDate = row.work_date;
    let dateVal = rawDate;
    if (typeof rawDate === 'string' && /^\d{4}-\d{2}-\d{2}/.test(rawDate)) {
      const [y, m, d] = rawDate.split('-').map(Number);
      dateVal = new Date(Date.UTC(y, m - 1, d));
    }

    const rowValues = [
      dateVal,
      row.line,
      row.unit,
      Number(row.operators_final) || 0,
      Number(row.reg_hours) || 8,
      Number(row.ot_man) || 0,
      Number(row.ot_hr) || 0,
      row.raw_buyer || '',
      row.buyer_group || '',
      row.style || '',
      String(row.item_no || ''), // text
      Number(row.order_qty) || 0,
      Number(row.smv) || 0,
      Number(row.earn_minutes) || 0,
      Number(row.production) || 0,
      Number(row.effi) || 0,
      Number(row.mc_hrs) || 0,
      '', // remarks left empty
    ];

    const excelRow = ws.addRow(rowValues);

    // Apply formatting per cell
    MERGED_COLUMNS.forEach((col, idx) => {
      const cell = excelRow.getCell(idx + 1);
      if (col.key === 'work_date') {
        cell.numFmt = 'dd-mmm-yyyy';
        cell.alignment = { horizontal: 'center' };
      } else if (col.numFmt) {
        cell.numFmt = col.numFmt;
      }

      if (col.key === 'item_no') {
        cell.numFmt = '@'; // plain text
      }
    });

    // Check for Red Flags
    const rawBuyerCell = excelRow.getCell(8); // Col 8: Raw Buyer
    const buyerCell = excelRow.getCell(9);    // Col 9: Buyer
    const opsCell = excelRow.getCell(4);      // Col 4: Operators

    if (row.is_pending || !row.buyer_group) {
      rawBuyerCell.fill = redFill;
      rawBuyerCell.font = redFont;
      rawBuyerCell.note = 'New buyer – needs mapping';

      buyerCell.fill = redFill;
      buyerCell.font = redFont;
      buyerCell.note = 'New buyer – needs mapping';
    }

    if (row.has_hp_pending) {
      opsCell.fill = redFill;
      opsCell.font = redFont;
      opsCell.note = 'HP present – decision needed';
    }
  }

  return wb.xlsx.writeBuffer();
}

/**
 * Writes the Final Summary Sheet (sheet name 'Sewtech') starting at cell D7.
 * Layout:
 *   D7: Sewtech
 *   D8: Buyer | E8: "14th Sep to 20th Sep" | F8: Avg prod/day | G8: Avg optrs /day
 *   D9..: Groups
 *   Last: Total (with Excel SUM formulas)
 *   Divisor note below
 */
async function writeFinalSummarySheet(summaryRows, { startDate, endDate, weekLabel, workingDays = 6 } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sewtech');

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };

  const headerDateLabel = formatOrdinalDateRange(startDate, endDate, weekLabel);

  // D7 title
  const cellD7 = ws.getCell('D7');
  cellD7.value = 'Sewtech';
  cellD7.font = { bold: true, size: 11 };

  // D8..G8 headers
  const headers = [
    { col: 'D', val: 'Buyer', align: 'left' },
    { col: 'E', val: headerDateLabel, align: 'right' },
    { col: 'F', val: 'Avg prod/day', align: 'right' },
    { col: 'G', val: 'Avg optrs /day', align: 'right' },
  ];

  headers.forEach((h) => {
    const cell = ws.getCell(`${h.col}8`);
    cell.value = h.val;
    cell.font = { bold: true };
    cell.border = thinBorder;
    cell.alignment = { vertical: 'middle', horizontal: h.align };
  });

  // Sort groups alphabetically
  const sorted = [...summaryRows].sort((a, b) => (a.buyer_group || '').localeCompare(b.buyer_group || ''));

  let startRow = 9;
  let endRow = startRow + sorted.length - 1;

  sorted.forEach((item, idx) => {
    const rIdx = startRow + idx;

    const cellD = ws.getCell(`D${rIdx}`);
    cellD.value = item.buyer_group;
    cellD.border = thinBorder;
    cellD.alignment = { vertical: 'middle', horizontal: 'left' };

    const cellE = ws.getCell(`E${rIdx}`);
    cellE.value = Number(item.sum_production) || 0;
    cellE.numFmt = '#,##0';
    cellE.border = thinBorder;
    cellE.alignment = { vertical: 'middle', horizontal: 'right' };

    const cellF = ws.getCell(`F${rIdx}`);
    cellF.value = Number(item.avg_prod_per_day) || (workingDays > 0 ? ((Number(item.sum_production) || 0) / workingDays) : 0);
    cellF.numFmt = '#,##0';
    cellF.border = thinBorder;
    cellF.alignment = { vertical: 'middle', horizontal: 'right' };

    const cellG = ws.getCell(`G${rIdx}`);
    cellG.value = Number(item.avg_optrs_per_day) || (workingDays > 0 ? ((Number(item.sum_operators) || 0) / workingDays) : 0);
    cellG.numFmt = '#,##0';
    cellG.border = thinBorder;
    cellG.alignment = { vertical: 'middle', horizontal: 'right' };
  });

  // Total row
  const totalRowIdx = endRow + 1;

  const cellDTotal = ws.getCell(`D${totalRowIdx}`);
  cellDTotal.value = 'Total';
  cellDTotal.font = { bold: true };
  cellDTotal.border = thinBorder;
  cellDTotal.alignment = { vertical: 'middle', horizontal: 'left' };

  const cellETotal = ws.getCell(`E${totalRowIdx}`);
  if (sorted.length > 0) {
    cellETotal.value = { formula: `SUM(E${startRow}:E${endRow})` };
  } else {
    cellETotal.value = 0;
  }
  cellETotal.font = { bold: true };
  cellETotal.numFmt = '#,##0';
  cellETotal.border = thinBorder;
  cellETotal.alignment = { vertical: 'middle', horizontal: 'right' };

  const cellFTotal = ws.getCell(`F${totalRowIdx}`);
  if (sorted.length > 0) {
    cellFTotal.value = { formula: `SUM(F${startRow}:F${endRow})` };
  } else {
    cellFTotal.value = 0;
  }
  cellFTotal.font = { bold: true };
  cellFTotal.numFmt = '#,##0';
  cellFTotal.border = thinBorder;
  cellFTotal.alignment = { vertical: 'middle', horizontal: 'right' };

  const cellGTotal = ws.getCell(`G${totalRowIdx}`);
  if (sorted.length > 0) {
    cellGTotal.value = { formula: `SUM(G${startRow}:G${endRow})` };
  } else {
    cellGTotal.value = 0;
  }
  cellGTotal.font = { bold: true };
  cellGTotal.numFmt = '#,##0';
  cellGTotal.border = thinBorder;
  cellGTotal.alignment = { vertical: 'middle', horizontal: 'right' };

  // Note row below Table 1
  const noteRowIdx = totalRowIdx + 2;
  const cellNote = ws.getCell(`D${noteRowIdx}`);
  cellNote.value = `Working days: ${workingDays}`;
  cellNote.font = { italic: true, size: 10, color: { argb: 'FF555555' } };

  // ==========================================
  // TABLE 2 — SEPARATE SEWTECH PRODUCTION & OPERATORS SUMMARY
  // ==========================================
  const t2HeaderRowIdx = noteRowIdx + 3; // e.g. Row 17

  const blueFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFCCE5FF' }, // Light blue header & total
  };

  const t2Headers = [
    { col: 'D', val: 'Row Labels', align: 'left' },
    { col: 'E', val: 'Sum of Production', align: 'right' },
    { col: 'F', val: 'Sum of Operators', align: 'right' },
  ];

  t2Headers.forEach((h) => {
    const cell = ws.getCell(`${h.col}${t2HeaderRowIdx}`);
    cell.value = h.val;
    cell.font = { bold: true };
    cell.fill = blueFill;
    cell.border = thinBorder;
    cell.alignment = { vertical: 'middle', horizontal: h.align };
  });

  const t2StartRow = t2HeaderRowIdx + 1;
  const t2EndRow = t2StartRow + sorted.length - 1;

  sorted.forEach((item, idx) => {
    const rIdx = t2StartRow + idx;

    const cellD = ws.getCell(`D${rIdx}`);
    cellD.value = item.buyer_group;
    cellD.border = thinBorder;
    cellD.alignment = { vertical: 'middle', horizontal: 'left' };

    const cellE = ws.getCell(`E${rIdx}`);
    cellE.value = Number(item.sum_production) || 0;
    cellE.numFmt = '#,##0';
    cellE.border = thinBorder;
    cellE.alignment = { vertical: 'middle', horizontal: 'right' };

    const cellF = ws.getCell(`F${rIdx}`);
    cellF.value = Number(item.sum_operators) || 0;
    cellF.numFmt = '#,##0';
    cellF.border = thinBorder;
    cellF.alignment = { vertical: 'middle', horizontal: 'right' };
  });

  // Table 2 Grand Total row
  const t2TotalRowIdx = t2EndRow + 1;

  const cellT2DTotal = ws.getCell(`D${t2TotalRowIdx}`);
  cellT2DTotal.value = 'Grand Total';
  cellT2DTotal.font = { bold: true };
  cellT2DTotal.fill = blueFill;
  cellT2DTotal.border = thinBorder;
  cellT2DTotal.alignment = { vertical: 'middle', horizontal: 'left' };

  const cellT2ETotal = ws.getCell(`E${t2TotalRowIdx}`);
  if (sorted.length > 0) {
    cellT2ETotal.value = { formula: `SUM(E${t2StartRow}:E${t2EndRow})` };
  } else {
    cellT2ETotal.value = 0;
  }
  cellT2ETotal.font = { bold: true };
  cellT2ETotal.fill = blueFill;
  cellT2ETotal.numFmt = '#,##0';
  cellT2ETotal.border = thinBorder;
  cellT2ETotal.alignment = { vertical: 'middle', horizontal: 'right' };

  const cellT2FTotal = ws.getCell(`F${t2TotalRowIdx}`);
  if (sorted.length > 0) {
    cellT2FTotal.value = { formula: `SUM(F${t2StartRow}:F${t2EndRow})` };
  } else {
    cellT2FTotal.value = 0;
  }
  cellT2FTotal.font = { bold: true };
  cellT2FTotal.fill = blueFill;
  cellT2FTotal.numFmt = '#,##0';
  cellT2FTotal.border = thinBorder;
  cellT2FTotal.alignment = { vertical: 'middle', horizontal: 'right' };

  // Set column widths
  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 4;
  ws.getColumn(3).width = 4;
  ws.getColumn(4).width = 24; // D: Buyer / Row Labels
  ws.getColumn(5).width = 26; // E: Production / Sum of Production
  ws.getColumn(6).width = 20; // F: Avg prod/day / Sum of Operators
  ws.getColumn(7).width = 18; // G: Avg optrs /day

  return wb.xlsx.writeBuffer();
}

/**
 * Parses a modified merged sheet re-uploaded by the user.
 */
function parseModifiedMergedSheet(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = wb.SheetNames.find((s) => s.toLowerCase() === 'sewtech') || wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

  if (!raw || raw.length < 2) {
    throw new Error('Modified sheet is empty or has no data rows.');
  }

  // Find header row containing 'Line' and 'Operators' or 'Buyer'
  let headerIdx = -1;
  for (let i = 0; i < Math.min(raw.length, 5); i++) {
    const r = (raw[i] || []).map((c) => String(c || '').trim().toLowerCase());
    if (r.includes('line') && (r.includes('buyer') || r.includes('operators'))) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx === -1) {
    throw new Error('Could not identify header row in modified sheet.');
  }

  const headers = raw[headerIdx].map((h) => String(h || '').trim());
  const dataRows = raw.slice(headerIdx + 1);

  const colMap = {};
  headers.forEach((h, idx) => {
    const low = h.toLowerCase();
    if (low === 'date') colMap.date = idx;
    else if (low === 'line') colMap.line = idx;
    else if (low === 'unit') colMap.unit = idx;
    else if (low === 'operators') colMap.operators_final = idx;
    else if (low.includes('reg') && low.includes('hour')) colMap.reg_hours = idx;
    else if (low.includes('ot') && low.includes('man')) colMap.ot_man = idx;
    else if (low.includes('ot') && low.includes('hr')) colMap.ot_hr = idx;
    else if (low === 'raw buyer') colMap.raw_buyer = idx;
    else if (low === 'buyer' && colMap.raw_buyer !== idx) colMap.buyer_group = idx;
    else if (low === 'style') colMap.style = idx;
    else if (low.includes('item')) colMap.item_no = idx;
    else if (low.includes('order') && low.includes('qty')) colMap.order_qty = idx;
    else if (low === 'smv') colMap.smv = idx;
    else if (low.includes('earn')) colMap.earn_minutes = idx;
    else if (low.includes('production')) colMap.production = idx;
    else if (low.includes('effi')) colMap.effi = idx;
    else if (low.includes('mc') && low.includes('hr')) colMap.mc_hrs = idx;
  });

  const parsedRows = [];
  for (let i = 0; i < dataRows.length; i++) {
    const r = dataRows[i];
    if (!r || r.every((c) => c === null || c === '')) continue;

    const lineVal = r[colMap.line];
    if (!lineVal || !/^LINE-\d+/i.test(String(lineVal).trim())) continue;

    const rowObj = {
      work_date: r[colMap.date],
      line: String(lineVal).trim(),
      unit: String(r[colMap.unit] || '').trim(),
      operators_final: Number(r[colMap.operators_final]) || 0,
      reg_hours: Number(r[colMap.reg_hours]) || 8,
      ot_man: Number(r[colMap.ot_man]) || 0,
      ot_hr: Number(r[colMap.ot_hr]) || 0,
      raw_buyer: String(r[colMap.raw_buyer] || '').trim(),
      buyer_group: String(r[colMap.buyer_group] || '').trim(),
      style: String(r[colMap.style] || '').trim(),
      item_no: String(r[colMap.item_no] || '').trim(),
      order_qty: Number(r[colMap.order_qty]) || 0,
      smv: Number(r[colMap.smv]) || 0,
      earn_minutes: Number(r[colMap.earn_minutes]) || 0,
      production: Number(r[colMap.production]) || 0,
      effi: Number(r[colMap.effi]) || 0,
      mc_hrs: Number(r[colMap.mc_hrs]) || 0,
    };

    parsedRows.push(rowObj);
  }

  return parsedRows;
}

module.exports = {
  formatOrdinalDateRange,
  writeMergedSheet,
  writeFinalSummarySheet,
  parseModifiedMergedSheet,
};
