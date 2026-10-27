/**
 * Reads the raw ERP export (.xls/.xlsx) into plain row objects, and writes the
 * cleaned "modification sheet" back out as .xlsx with new-brand rows highlighted red —
 * exactly mirroring what your manager does by hand.
 */
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

// Column header -> internal field name (matches the header row found in your real files)
const COLUMN_MAP = {
  'Facility': 'facility',
  'PCH': 'pch',
  'Act Ctrl Obj': 'act_ctrl_obj',
  'Sales Person': 'sales_person',
  'Item Responsible': 'item_responsible',
  'Item': 'item',
  'Item Group': 'item_group',
  'Type': 'type',
  'Item Category': 'item_category',
  'Shahi Style': 'shahi_style',
  'Unit': 'unit',
  'Brand': 'brand',
  'Buyer': 'raw_buyer',
  'Order Qty': 'order_qty',
  'Production Qty': 'production_qty',
  'Budgeted Cost': 'budgeted_cost',
  'Actual Cost': 'actual_cost',
  'OT Cost': 'ot_cost',
  'Total Cost': 'total_cost',
  'Cost/Pcs': 'cost_per_pcs',
  'Sam': 'sam',
  'Budgeted Cost/SAM': 'budgeted_cost_per_sam',
  'Actual Cost/SAM': 'actual_cost_per_sam',
  'Budgeted Efficiency': 'budgeted_efficiency',
  'Actual Efficiency': 'actual_efficiency',
  'FOB/GMT\nin INR': 'fob_gmt_inr',
  'Currency': 'currency',
  'FOB/GMT\nin FC': 'fob_gmt_fc',
  'Currency\nConversion': 'currency_conversion',
};

function findHeaderRow(sheetRows) {
  for (let i = 0; i < Math.min(sheetRows.length, 10); i++) {
    const row = sheetRows[i] || [];
    if (row.includes('Sales Person') && row.includes('Brand')) {
      return i;
    }
  }
  throw new Error('Could not locate header row (expected "Sales Person" and "Brand" columns) in first 10 rows.');
}

function parseRawWorkbook(buffer, filename) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

  const headerIdx = findHeaderRow(raw);
  const headers = raw[headerIdx];
  const dataRows = raw.slice(headerIdx + 1);

  const fieldByCol = headers.map((h) => COLUMN_MAP[h] || null);

  const rows = [];
  for (let idx = 0; idx < dataRows.length; idx++) {
    const r = dataRows[idx];
    // skip fully blank trailing rows
    if (!r || r.every((v) => v === null || v === '')) continue;
    const obj = {};
    fieldByCol.forEach((field, i) => {
      if (field) obj[field] = r[i] === undefined ? null : r[i];
    });
    // A row with no Brand and no Item is not real data (e.g. footer/summary rows)
    if (obj.brand == null && obj.item == null) continue;
    obj.raw_row_no = headerIdx + 2 + idx;
    rows.push(obj);
  }
  return rows;
}

const NUMERIC_FIELDS = [
  'order_qty', 'production_qty', 'budgeted_cost', 'actual_cost', 'ot_cost', 'total_cost',
  'cost_per_pcs', 'sam', 'budgeted_cost_per_sam', 'actual_cost_per_sam',
  'budgeted_efficiency', 'actual_efficiency', 'fob_gmt_inr', 'fob_gmt_fc', 'currency_conversion',
];

function coerceNumbers(rows) {
  for (const row of rows) {
    for (const f of NUMERIC_FIELDS) {
      if (row[f] != null && row[f] !== '') row[f] = Number(row[f]);
    }
  }
}

const EXPORT_COLUMNS = [
  { header: 'Facility', key: 'facility', type: 'text' },
  { header: 'PCH', key: 'pch', type: 'text' },
  { header: 'Act Ctrl Obj', key: 'act_ctrl_obj', type: 'text' },
  { header: 'Sales Person', key: 'sales_person', type: 'text' },
  { header: 'Item Responsible', key: 'item_responsible', type: 'text' },
  { header: 'Item', key: 'item', type: 'text' },
  { header: 'Item Group', key: 'item_group', type: 'text' },
  { header: 'Type', key: 'type', type: 'text' },
  { header: 'Item Category', key: 'item_category', type: 'text' },
  { header: 'Shahi Style', key: 'shahi_style', type: 'text' },
  { header: 'Unit', key: 'unit', type: 'text' },
  { header: 'Brand', key: 'brand', type: 'text' },
  { header: 'Buyer', key: 'resolved_buyer', type: 'text' },
  { header: 'Order Qty', key: 'order_qty', type: 'number', numFmt: undefined },
  { header: 'Production Qty', key: 'production_qty', type: 'number', numFmt: '#,##0' },
  { header: 'Mcs', key: 'mcs', type: 'number', numFmt: '#,##0' },
  { header: 'Budgeted Cost', key: 'budgeted_cost', type: 'number', numFmt: '#,#00.##' },
  { header: 'Actual Cost', key: 'actual_cost', type: 'number', numFmt: '#,#00.##' },
  { header: 'OT Cost', key: 'ot_cost', type: 'number', numFmt: '#,#00.##' },
  { header: 'Total Cost', key: 'total_cost', type: 'number', numFmt: '#,#00.##' },
  { header: 'Cost/Pcs', key: 'cost_per_pcs', type: 'number', numFmt: '#0.##' },
  { header: 'Sam', key: 'sam', type: 'number', numFmt: '#0.##' },
  { header: 'Budgeted Cost/SAM', key: 'budgeted_cost_per_sam', type: 'number', numFmt: '#0.##' },
  { header: 'Actual Cost/SAM', key: 'actual_cost_per_sam', type: 'number', numFmt: '#0.00;#0.00\\-' },
  { header: 'Budgeted Efficiency', key: 'budgeted_efficiency', type: 'number', numFmt: '#0' },
  { header: 'Actual Efficiency', key: 'actual_efficiency', type: 'number', numFmt: '#0' },
  { header: 'FOB/GMT in INR', key: 'fob_gmt_inr', type: 'number', numFmt: '#0.00' },
  { header: 'Currency', key: 'currency', type: 'text' },
  { header: 'FOB/GMT in FC', key: 'fob_gmt_fc', type: 'number', numFmt: '#0.00' },
  { header: 'Currency Conversion', key: 'currency_conversion', type: 'number', numFmt: '#0.00' },
];

/**
 * Writes the cleaned working sheet, with rows for never-seen brands highlighted red
 * (same signal your manager currently relies on manual review for).
 */
async function writeModificationSheet(rows, { weekLabel } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('itemAndUnitWiseCostBreakUp');

  ws.addRow(['Item And Unit Wise Cost Break Up']);
  ws.addRow([]);
  ws.addRow(EXPORT_COLUMNS.map((col) => col.header));
  const headerRow = ws.getRow(3);
  headerRow.font = { bold: true };

  const redFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
  const redFont = { color: { argb: 'FF9C0006' }, bold: true };

  rows.filter((r) => !r.was_deleted).forEach((row) => {
    const values = EXPORT_COLUMNS.map((col) => {
      const rawVal = row[col.key];
      if (rawVal === null || rawVal === undefined || rawVal === '') {
        return null;
      }
      if (col.type === 'number') {
        const num = Number(rawVal);
        return isNaN(num) ? rawVal : num;
      }
      return rawVal;
    });

    const excelRow = ws.addRow(values);

    EXPORT_COLUMNS.forEach((col, idx) => {
      const cell = excelRow.getCell(idx + 1);
      if (col.numFmt) {
        cell.numFmt = col.numFmt;
      }
      if (row.is_new_brand) {
        cell.fill = redFill;
        cell.font = redFont;
      }
    });
  });

  ws.columns.forEach((col) => { col.width = 16; });

  return wb.xlsx.writeBuffer();
}

async function writeFinalSummarySheet(pivotRows, { weekLabel, workingDays = 5 } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Final Summary');

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };

  // Top two lines with bold text
  const cellA1 = ws.getCell('A1');
  cellA1.value = `Week: ${weekLabel || '-'}`;
  cellA1.font = { bold: true, size: 11 };

  const cellA2 = ws.getCell('A2');
  cellA2.value = `Working days: ${workingDays}`;
  cellA2.font = { bold: true, size: 11 };

  // 1. Prepare data
  const leftRows = [...pivotRows].sort((a, b) => (a.buyer_group || '').localeCompare(b.buyer_group || ''));
  const rightRows = [...pivotRows].sort((a, b) => (Number(b.sum_production_qty) || 0) - (Number(a.sum_production_qty) || 0));

  let rawTotalProdQty = 0;
  let rawTotalAvgProd = 0;
  let rawTotalAvgOptrs = 0;
  let rawTotalMcs = 0;

  for (const g of pivotRows) {
    const rawProdQty = Number(g.sum_production_qty) || 0;
    const rawAvgProd = g.avg_prod_per_day != null && g.avg_prod_per_day !== '' ? Number(g.avg_prod_per_day) : (rawProdQty / workingDays);
    const rawAvgOptrs = Number(g.avg_optrs_per_day);
    const validAvgOptrs = Number.isFinite(rawAvgOptrs) ? rawAvgOptrs : 0;
    const rawMcs = Number(g.sum_mcs) || 0;

    rawTotalProdQty += rawProdQty;
    rawTotalAvgProd += (Number.isFinite(rawAvgProd) ? rawAvgProd : 0);
    rawTotalAvgOptrs += validAvgOptrs;
    rawTotalMcs += rawMcs;
  }

  // --- Row 4 ---
  // Left table: B4:C4 merged with "Data"
  ws.mergeCells('B4:C4');
  const dataCellB4 = ws.getCell('B4');
  dataCellB4.value = 'Data';
  dataCellB4.alignment = { vertical: 'middle', horizontal: 'left' };
  dataCellB4.border = thinBorder;
  ws.getCell('C4').border = thinBorder;

  // --- Row 5 (Headers) ---
  // Left Table Headers (Cols A, B, C)
  const cellA5 = ws.getCell('A5');
  cellA5.value = 'Buyer';
  cellA5.alignment = { vertical: 'middle', horizontal: 'left' };
  cellA5.border = thinBorder;

  const cellB5 = ws.getCell('B5');
  cellB5.value = 'Sum of Production Qty';
  cellB5.alignment = { vertical: 'middle', horizontal: 'right' };
  cellB5.border = thinBorder;

  const cellC5 = ws.getCell('C5');
  cellC5.value = 'Sum of Mcs';
  cellC5.alignment = { vertical: 'middle', horizontal: 'right' };
  cellC5.border = thinBorder;

  // Right Table Headers (Cols F, G, H, I) - Cols D, E remain blank
  const headerLabel = weekLabel || 'Production Qty';

  const cellF5 = ws.getCell('F5');
  cellF5.value = 'Buyer';
  cellF5.font = { bold: true };
  cellF5.alignment = { vertical: 'middle', horizontal: 'center' };
  cellF5.border = thinBorder;

  const cellG5 = ws.getCell('G5');
  cellG5.value = headerLabel;
  cellG5.font = { bold: true };
  cellG5.alignment = { vertical: 'middle', horizontal: 'center' };
  cellG5.border = thinBorder;

  const cellH5 = ws.getCell('H5');
  cellH5.value = 'Avg prod /day';
  cellH5.font = { bold: true };
  cellH5.alignment = { vertical: 'middle', horizontal: 'center' };
  cellH5.border = thinBorder;

  const cellI5 = ws.getCell('I5');
  cellI5.value = 'Avg optrs /day';
  cellI5.font = { bold: true };
  cellI5.alignment = { vertical: 'middle', horizontal: 'center' };
  cellI5.border = thinBorder;

  // --- Data Rows (Row 6 onwards) ---
  const maxRows = Math.max(leftRows.length, rightRows.length);
  for (let i = 0; i < maxRows; i++) {
    const rowIdx = i + 6;

    // Left table item
    if (i < leftRows.length) {
      const lg = leftRows[i];
      const prodQty = Math.round(Number(lg.sum_production_qty) || 0);
      const mcs = Math.round(Number(lg.sum_mcs) || 0);

      const cellA = ws.getCell(`A${rowIdx}`);
      cellA.value = lg.buyer_group;
      cellA.alignment = { vertical: 'middle', horizontal: 'left' };
      cellA.border = thinBorder;

      const cellB = ws.getCell(`B${rowIdx}`);
      cellB.value = prodQty;
      cellB.numFmt = '#,##0';
      cellB.alignment = { vertical: 'middle', horizontal: 'right' };
      cellB.border = thinBorder;

      const cellC = ws.getCell(`C${rowIdx}`);
      cellC.value = mcs;
      cellC.numFmt = '#,##0';
      cellC.alignment = { vertical: 'middle', horizontal: 'right' };
      cellC.border = thinBorder;
    }

    // Right table item
    if (i < rightRows.length) {
      const rg = rightRows[i];
      const rawProdQty = Number(rg.sum_production_qty) || 0;
      const rawAvgProd = rg.avg_prod_per_day != null && rg.avg_prod_per_day !== '' ? Number(rg.avg_prod_per_day) : (rawProdQty / workingDays);
      const rawAvgOptrs = Number(rg.avg_optrs_per_day);
      const validAvgOptrs = Number.isFinite(rawAvgOptrs) ? rawAvgOptrs : 0;

      const prodQty = Math.round(rawProdQty);
      const avgProd = Math.round(rawAvgProd);
      const avgOptrs = Math.round(validAvgOptrs);

      const cellF = ws.getCell(`F${rowIdx}`);
      cellF.value = rg.buyer_group;
      cellF.alignment = { vertical: 'middle', horizontal: 'center' };
      cellF.border = thinBorder;

      const cellG = ws.getCell(`G${rowIdx}`);
      cellG.value = prodQty;
      cellG.numFmt = '#,##0';
      cellG.alignment = { vertical: 'middle', horizontal: 'center' };
      cellG.border = thinBorder;

      const cellH = ws.getCell(`H${rowIdx}`);
      cellH.value = avgProd;
      cellH.numFmt = '#,##0';
      cellH.alignment = { vertical: 'middle', horizontal: 'center' };
      cellH.border = thinBorder;

      const cellI = ws.getCell(`I${rowIdx}`);
      cellI.value = avgOptrs;
      cellI.numFmt = '#,##0';
      cellI.alignment = { vertical: 'middle', horizontal: 'center' };
      cellI.border = thinBorder;
    }
  }

  // --- Total Rows ---
  const totalRowIdx = maxRows + 6;

  // Left Table Grand Total
  const cellATotal = ws.getCell(`A${totalRowIdx}`);
  cellATotal.value = 'Grand Total';
  cellATotal.font = { bold: true };
  cellATotal.alignment = { vertical: 'middle', horizontal: 'left' };
  cellATotal.border = thinBorder;

  const cellBTotal = ws.getCell(`B${totalRowIdx}`);
  cellBTotal.value = Math.round(rawTotalProdQty);
  cellBTotal.font = { bold: true };
  cellBTotal.numFmt = '#,##0';
  cellBTotal.alignment = { vertical: 'middle', horizontal: 'right' };
  cellBTotal.border = thinBorder;

  const cellCTotal = ws.getCell(`C${totalRowIdx}`);
  cellCTotal.value = Math.round(rawTotalMcs);
  cellCTotal.font = { bold: true };
  cellCTotal.numFmt = '#,##0';
  cellCTotal.alignment = { vertical: 'middle', horizontal: 'right' };
  cellCTotal.border = thinBorder;

  // Right Table Total
  const cellFTotal = ws.getCell(`F${totalRowIdx}`);
  cellFTotal.value = 'Total';
  cellFTotal.font = { bold: true };
  cellFTotal.alignment = { vertical: 'middle', horizontal: 'center' };
  cellFTotal.border = thinBorder;

  const cellGTotal = ws.getCell(`G${totalRowIdx}`);
  cellGTotal.value = Math.round(rawTotalProdQty);
  cellGTotal.font = { bold: true };
  cellGTotal.numFmt = '#,##0';
  cellGTotal.alignment = { vertical: 'middle', horizontal: 'center' };
  cellGTotal.border = thinBorder;

  const cellHTotal = ws.getCell(`H${totalRowIdx}`);
  cellHTotal.value = Math.round(rawTotalAvgProd);
  cellHTotal.font = { bold: true };
  cellHTotal.numFmt = '#,##0';
  cellHTotal.alignment = { vertical: 'middle', horizontal: 'center' };
  cellHTotal.border = thinBorder;

  const cellITotal = ws.getCell(`I${totalRowIdx}`);
  cellITotal.value = Math.round(rawTotalAvgOptrs);
  cellITotal.font = { bold: true };
  cellITotal.numFmt = '#,##0';
  cellITotal.alignment = { vertical: 'middle', horizontal: 'center' };
  cellITotal.border = thinBorder;

  // Column widths
  ws.getColumn(1).width = 22; // Col A
  ws.getColumn(2).width = 24; // Col B
  ws.getColumn(3).width = 16; // Col C
  ws.getColumn(4).width = 5;  // Col D (blank gap)
  ws.getColumn(5).width = 5;  // Col E (blank gap)
  ws.getColumn(6).width = 22; // Col F
  ws.getColumn(7).width = 24; // Col G
  ws.getColumn(8).width = 18; // Col H
  ws.getColumn(9).width = 18; // Col I

  return wb.xlsx.writeBuffer();
}

module.exports = {
  parseRawWorkbook,
  coerceNumbers,
  writeModificationSheet,
  writeFinalSummarySheet,
};
