const XLSX = require('xlsx');

function normalizedHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeText(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function getColumn(row, name) {
  const wanted = normalizedHeader(name);
  const key = Object.keys(row).find((candidate) => normalizedHeader(candidate) === wanted);
  return key === undefined ? '' : row[key];
}

function normalizeExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value);
    return date ? { year: date.y, month: date.m, day: date.d } : null;
  }

  const valueText = normalizeText(value);
  if (!valueText) return null;

  let match = valueText.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2}|\d{4})$/);
  if (match) {
    const year = match[3].length === 2 ? 2000 + Number(match[3]) : Number(match[3]);
    return { year, month: Number(match[2]), day: Number(match[1]) };
  }

  match = valueText.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (match) return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };

  match = valueText.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (match) return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3]) };

  return null;
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const valueText = normalizeText(value).replace(/,/g, '');
  if (!valueText) return 0;
  const parsed = Number(valueText);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatMonthLabel(dateObj) {
  if (!dateObj) return '';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[dateObj.month - 1] || '';
  const year = String(dateObj.year).slice(-2);
  return `${month}-${year}`;
}

function parseChartFile(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false, raw: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('Excel file has no sheets.');

  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true });
  if (matrix.length < 2) throw new Error('Excel file must have header row.');

  // Find header row by looking for "Season" or other key headers
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(5, matrix.length); i++) {
    const row = matrix[i];
    const headerText = row.map((cell) => normalizedHeader(cell)).join('|');
    if (headerText.includes('season') && headerText.includes('style')) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    throw new Error(
      'Could not find header row with "Season" and "Style" columns. Please check the file structure.',
    );
  }

  const headers = matrix[headerRowIndex].map((header) => normalizeText(header));

  // Validate required columns
  const requiredColumns = ['Season', 'Style', 'Total Qty', 'Proj xfty'];
  const missingColumns = [];

  for (const col of requiredColumns) {
    const found = headers.some((h) => normalizedHeader(h) === normalizedHeader(col));
    if (!found) missingColumns.push(col);
  }

  if (missingColumns.length > 0) {
    throw new Error(
      `Required column(s) not found: ${missingColumns.join(', ')}. ` +
      `File must contain columns: Season, Style, Total Qty, Proj xfty`,
    );
  }

  const rows = [];
  for (let i = headerRowIndex + 1; i < matrix.length; i++) {
    const values = matrix[i];
    const row = {};
    headers.forEach((header, columnIndex) => {
      if (header) row[header] = values[columnIndex] ?? '';
    });

    // Only include rows with data
    if (Object.keys(row).some((key) => normalizeText(row[key]) !== '')) {
      rows.push(row);
    }
  }

  return rows;
}

function extractWeekCodeFromFilename(filename) {
  // Extract week code from filename like "WK-202633.xlsx" or "WK202634.xlsx"
  const match = filename.match(/WK-?(\d{6})/i);
  if (match) {
    const digits = match[1];
    const weekNum = digits.slice(-2);
    return `wk-${weekNum}`;
  }
  return null;
}

function aggregateChartData(rows) {
  const aggregated = new Map(); // key = "Season|Style|MonthLabel"
  const monthSet = new Set();

  for (const row of rows) {
    const season = normalizeText(getColumn(row, 'Season'));
    const style = normalizeText(getColumn(row, 'Style'));
    const totalQty = toNumber(getColumn(row, 'Total Qty'));
    const projXfty = getColumn(row, 'Proj xfty');

    if (!season || !style) continue;

    const dateObj = normalizeExcelDate(projXfty);
    if (!dateObj) continue;

    const monthLabel = formatMonthLabel(dateObj);
    monthSet.add({ label: monthLabel, date: dateObj });

    const key = `${season}|${style}|${monthLabel}`;
    const current = aggregated.get(key) || 0;
    aggregated.set(key, current + totalQty);
  }

  return { aggregated, months: Array.from(monthSet).sort((a, b) => {
    if (a.date.year !== b.date.year) return a.date.year - b.date.year;
    return a.date.month - b.date.month;
  }) };
}

function getMonthRange(months, startMonth, endMonth) {
  const candidateMonths = Array.isArray(months) ? months : [];
  if (!candidateMonths.length) return candidateMonths;

  if (!startMonth && !endMonth) return candidateMonths;

  const startIdx = startMonth ? candidateMonths.findIndex((month) => month.label === startMonth) : 0;
  const endIdx = endMonth ? candidateMonths.findIndex((month) => month.label === endMonth) : candidateMonths.length - 1;

  if (startIdx === -1 || endIdx === -1) {
    throw new Error(`Selected month range is invalid. Available months: ${candidateMonths.map((m) => m.label).join(', ')}`);
  }

  if (startIdx > endIdx) {
    throw new Error('Start Month cannot be later than End Month.');
  }

  return candidateMonths.slice(startIdx, endIdx + 1);
}

function buildSeasonStyleMap(aggregated) {
  const map = new Map();
  for (const key of aggregated.keys()) {
    const [season, style] = key.split('|');
    const seasonKey = season;
    if (!map.has(seasonKey)) map.set(seasonKey, new Set());
    map.get(seasonKey).add(style);
  }
  return map;
}

function generateChartSheet(weekCode, aggregated, months, sheetName) {
  const seasonStyleMap = buildSeasonStyleMap(aggregated);
  const sortedSeasons = Array.from(seasonStyleMap.keys()).sort();

  const monthLabels = months.map((m) => m.label);
  const headerRow = ['Chart', 'Season', 'Style', ...monthLabels, 'Grand Total'];

  const rows = [
    new Array(headerRow.length).fill(''), // Blank row for spacing
    headerRow,
  ];

  let grandTotal = 0;
  const grandMonthTotals = Object.fromEntries(monthLabels.map((label) => [label, 0]));

  for (const season of sortedSeasons) {
    const styles = Array.from(seasonStyleMap.get(season)).sort();
    let seasonTotal = 0;
    const seasonMonthTotals = Object.fromEntries(monthLabels.map((label) => [label, 0]));

    for (const style of styles) {
      const rowData = ['', season, style];
      let styleGrandTotal = 0;

      for (const month of months) {
        const key = `${season}|${style}|${month.label}`;
        const qty = aggregated.get(key) || 0;
        rowData.push(qty);
        styleGrandTotal += qty;
        seasonMonthTotals[month.label] += qty;
        grandMonthTotals[month.label] += qty;
      }

      // FIX (bug #3): skip styles with zero volume across the selected month
      // range. Without this, every style name that exists anywhere in the raw
      // file gets a row, even ones with no data in the displayed window.
      if (styleGrandTotal === 0) continue;

      rowData.push(styleGrandTotal);
      rows.push(rowData);
      seasonTotal += styleGrandTotal;
    }

    // Season total row
    const seasonTotalRow = ['', `${season} Total`, '', ...monthLabels.map((label) => seasonMonthTotals[label])];
    seasonTotalRow[3 + monthLabels.length] = seasonTotal;
    rows.push(seasonTotalRow);

    grandTotal += seasonTotal;
  }

  // Grand Total row
  const grandTotalRow = ['', 'Grand Total', '', ...monthLabels.map((label) => grandMonthTotals[label])];
  grandTotalRow[3 + monthLabels.length] = grandTotal;
  rows.push(grandTotalRow);

  // Create worksheet
  const worksheet = XLSX.utils.aoa_to_sheet(rows);

  // Set column widths
  const colWidths = [8, 12, 20, ...monthLabels.map(() => 12), 12];
  worksheet['!cols'] = colWidths.map((w) => ({ wch: w }));

  // Freeze header rows
  worksheet['!freeze'] = { xSplit: 0, ySplit: 2 };

  // Format header row (row 2, 0-indexed as row 1)
  const headerRowIndex = 1;
  for (let col = 0; col < headerRow.length; col++) {
    const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
    if (!worksheet[cellAddress]) worksheet[cellAddress] = {};
    worksheet[cellAddress].fill = { type: 'solid', fgColor: { rgb: 'FFFF00' } }; // Yellow
    worksheet[cellAddress].font = { bold: true };
  }

  return worksheet;
}

function getWorksheetMonthLabels(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, blankrows: false });

  if (!rows.length) return [];

  const firstRow = rows[0] || [];
  if (Array.isArray(firstRow) && firstRow.some((value) => String(value || '').includes('Sum of '))) {
    const labels = [];
    for (let col = 2; col < firstRow.length; col += 3) {
      const value = firstRow[col];
      if (value && /^Sum of /i.test(String(value))) {
        labels.push(String(value).replace(/^Sum of /i, '').trim());
      }
    }
    return labels;
  }

  const secondRow = rows[1] || [];
  return secondRow.slice(3, -1).filter((value) => value && String(value).trim() !== 'Grand Total');
}

function validateWorkbookMonthRange(workbook, expectedMonths) {
  const expected = expectedMonths.map((month) => String(month));
  const entries = workbook.SheetNames.map((sheetName) => ({
    name: sheetName,
    months: getWorksheetMonthLabels(workbook.Sheets[sheetName]),
  }));

  for (const entry of entries) {
    if (entry.months.length !== expected.length || entry.months.some((month, idx) => month !== expected[idx])) {
      const details = entries.map((sheet) => `${sheet.name}: ${sheet.months.join(' -> ') || 'none'}`).join('\n');
      throw new Error(`Month range mismatch detected.\nExpected: ${expected.join(' -> ')}\n${details}\nWorkbook generation aborted.`);
    }
  }
}

function generateComparisonSheet(previousAgg, currentAgg, months, previousWeekCode, currentWeekCode) {
  const allKeys = new Set();
  for (const key of previousAgg.keys()) {
    const [season, style] = key.split('|');
    allKeys.add(`${season}|${style}`);
  }
  for (const key of currentAgg.keys()) {
    const [season, style] = key.split('|');
    allKeys.add(`${season}|${style}`);
  }

  const sortedKeys = Array.from(allKeys).sort();
  const seasonStyleMap = new Map();
  for (const key of sortedKeys) {
    const [season, style] = key.split('|');
    if (!seasonStyleMap.has(season)) seasonStyleMap.set(season, []);
    seasonStyleMap.get(season).push(style);
  }

  const monthLabels = months.map((m) => m.label);

  // Row 1: Month group headers (merged headers)
  const row1 = ['', '']; // Empty for Season and Style columns
  for (const month of monthLabels) {
    row1.push(`Sum of ${month}`, '', '');
  }
  // FIX (bug #4): Total Sum needs 3 cells (wk-33 / wk-34 / Variation), same
  // shape as every month block, not 2.
  row1.push('Total Sum', '', '');

  // Row 2: Sub-headers
  const row2 = ['Season', 'Style'];
  for (const month of monthLabels) {
    row2.push(previousWeekCode, currentWeekCode, 'Variation');
  }
  // FIX (bug #4): add the missing 'Variation' sub-header for the total block.
  row2.push(previousWeekCode, currentWeekCode, 'Variation');

  // Data rows
  const dataRows = [];
  let grandTotalPrev = 0;
  let grandTotalCurr = 0;

  const sortedSeasons = Array.from(seasonStyleMap.keys()).sort();

  let currentRowIndex = 2; // Start after headers

  for (const season of sortedSeasons) {
    const styles = seasonStyleMap.get(season);
    const seasonMonthTotals = {}; // { monthLabel: { prev, curr } }

    // Initialize month totals for this season
    for (const month of monthLabels) {
      seasonMonthTotals[month] = { prev: 0, curr: 0 };
    }

    // Process each style in this season
    for (const style of styles) {
      const rowData = [season, style];
      let styleTotalPrev = 0;
      let styleTotalCurr = 0;

      for (const month of monthLabels) {
        const prevKey = `${season}|${style}|${month}`;
        const currKey = `${season}|${style}|${month}`;
        const prevQty = previousAgg.get(prevKey) || 0;
        const currQty = currentAgg.get(currKey) || 0;
        const variation = currQty - prevQty;

        rowData.push(prevQty, currQty, variation);
        styleTotalPrev += prevQty;
        styleTotalCurr += currQty;

        // Accumulate for season totals (always, even if this style ends up
        // filtered out below — a style with zero volume contributes zero
        // either way, so this keeps the season totals correct regardless).
        seasonMonthTotals[month].prev += prevQty;
        seasonMonthTotals[month].curr += currQty;
      }

      // FIX (bug #3): skip styles with no volume in either week across the
      // selected month range, instead of writing an all-zero row for every
      // style name that exists anywhere in the raw files.
      if (styleTotalPrev === 0 && styleTotalCurr === 0) continue;

      // Style total sum (across all months) + total variation
      rowData.push(styleTotalPrev, styleTotalCurr, styleTotalCurr - styleTotalPrev);
      rowData._isDataRow = true;
      dataRows.push(rowData);
      currentRowIndex++;
    }

    // Season total row
    const seasonTotalRow = [`${season} Total`, ''];
    let seasonTotalPrev = 0;
    let seasonTotalCurr = 0;

    for (const month of monthLabels) {
      const prevMonthTotal = seasonMonthTotals[month].prev;
      const currMonthTotal = seasonMonthTotals[month].curr;
      const variation = currMonthTotal - prevMonthTotal;

      seasonTotalRow.push(prevMonthTotal, currMonthTotal, variation);
      seasonTotalPrev += prevMonthTotal;
      seasonTotalCurr += currMonthTotal;
    }

    seasonTotalRow.push(seasonTotalPrev, seasonTotalCurr, seasonTotalCurr - seasonTotalPrev);
    seasonTotalRow._isSeasonTotal = true;
    seasonTotalRow._rowIndex = currentRowIndex;
    dataRows.push(seasonTotalRow);
    currentRowIndex++;

    grandTotalPrev += seasonTotalPrev;
    grandTotalCurr += seasonTotalCurr;
  }

  // Grand Total row
  const grandTotalRow = ['Grand Total', ''];
  const grandMonthTotals = {}; // { monthLabel: { prev, curr } }

  // Initialize
  for (const month of monthLabels) {
    grandMonthTotals[month] = { prev: 0, curr: 0 };
  }

  // Recalculate from data rows to get month-level grand totals
  for (const row of dataRows) {
    if (row._isSeasonTotal) {
      for (let i = 0; i < monthLabels.length; i++) {
        const colIdx = 2 + i * 3;
        grandMonthTotals[monthLabels[i]].prev += row[colIdx] || 0;
        grandMonthTotals[monthLabels[i]].curr += row[colIdx + 1] || 0;
      }
    }
  }

  for (const month of monthLabels) {
    const prevMonthTotal = grandMonthTotals[month].prev;
    const currMonthTotal = grandMonthTotals[month].curr;
    const variation = currMonthTotal - prevMonthTotal;

    grandTotalRow.push(prevMonthTotal, currMonthTotal, variation);
  }

  grandTotalRow.push(grandTotalPrev, grandTotalCurr, grandTotalCurr - grandTotalPrev);
  grandTotalRow._isGrandTotal = true;
  grandTotalRow._rowIndex = currentRowIndex;
  dataRows.push(grandTotalRow);

  // Combine headers and data
  const allRows = [row1, row2, ...dataRows];

  const worksheet = XLSX.utils.aoa_to_sheet(allRows);

  // Set every monthly Variation cell to a live formula:
  // Aug = Dx-Cx, Sep = Gx-Fx, Oct = Jx-Ix, Nov = Mx-Lx ...
  // (current-week column minus previous-week column, matching the manual
  // chart's own E-D convention, and matching the `variation` value computed
  // above.)
  for (let rowIdx = 2; rowIdx < allRows.length; rowIdx++) {
    for (let monthIdx = 0; monthIdx < monthLabels.length; monthIdx++) {
      const prevCol = 2 + monthIdx * 3;
      const currCol = prevCol + 1;
      const varCol = prevCol + 2;
      const prevCell = XLSX.utils.encode_cell({ r: rowIdx, c: prevCol });
      const currCell = XLSX.utils.encode_cell({ r: rowIdx, c: currCol });
      const varCell = XLSX.utils.encode_cell({ r: rowIdx, c: varCol });

      worksheet[varCell] = {
        ...worksheet[varCell],
        // FIX (bug #1): no leading '=' here — the xlsx library adds it when
        // writing the file. Including one produces a literal "==" formula.
        // FIX (bug #2): currCell - prevCell (not prevCell - currCell), to
        // match the manual chart's wk34-wk33 convention.
        f: `${currCell}-${prevCell}`,
        t: 'n',
      };
      delete worksheet[varCell].v;
    }
  }

  // FIX (bug #4): write the Total Sum block's Variation formula too — this
  // block was previously skipped entirely.
  const totalPrevCol = 2 + monthLabels.length * 3;
  const totalCurrCol = totalPrevCol + 1;
  const totalVarCol = totalPrevCol + 2;
  for (let rowIdx = 2; rowIdx < allRows.length; rowIdx++) {
    const totalPrevCell = XLSX.utils.encode_cell({ r: rowIdx, c: totalPrevCol });
    const totalCurrCell = XLSX.utils.encode_cell({ r: rowIdx, c: totalCurrCol });
    const totalVarCell = XLSX.utils.encode_cell({ r: rowIdx, c: totalVarCol });

    worksheet[totalVarCell] = {
      ...worksheet[totalVarCell],
      f: `${totalCurrCell}-${totalPrevCell}`,
      t: 'n',
    };
    delete worksheet[totalVarCell].v;
  }

  // Set column widths - increased for better readability
  const colWidths = [14, 32, ...monthLabels.flatMap(() => [12, 12, 12]), 12, 12, 12];
  worksheet['!cols'] = colWidths.map((w) => ({ wch: w }));

  // Set row heights
  worksheet['!rows'] = [];
  for (let r = 0; r < allRows.length; r++) {
    worksheet['!rows'][r] = { hpt: r === 0 || r === 1 ? 18 : 16 }; // Larger headers
  }

  // Freeze header rows
  worksheet['!freeze'] = { xSplit: 2, ySplit: 2 };

  // Helper function: Apply border to cell
  function applyBorder(cellAddress, borderStyle) {
    if (!worksheet[cellAddress]) worksheet[cellAddress] = {};
    worksheet[cellAddress].border = borderStyle;
  }

  // Helper function: Apply formatting
  function applyCellFormat(cellAddress, format) {
    if (!worksheet[cellAddress]) worksheet[cellAddress] = {};
    Object.assign(worksheet[cellAddress], format);
  }

  // Define border styles
  const thinBorder = {
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
  };

  const thickBorder = {
    left: { style: 'thick', color: { rgb: '000000' } },
    right: { style: 'thick', color: { rgb: '000000' } },
    top: { style: 'thick', color: { rgb: '000000' } },
    bottom: { style: 'thick', color: { rgb: '000000' } },
  };

  const thickTopBorder = {
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
    top: { style: 'thick', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
  };

  // Format Row 1: Month headers (merged)
  const monthHeaderFormat = {
    fill: { type: 'solid', fgColor: { rgb: 'FFD966' } }, // Bright yellow
    font: { bold: true, size: 11 },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: thickBorder,
  };

  const emptyHeaderFormat = {
    fill: { type: 'solid', fgColor: { rgb: 'FFD966' } },
    border: thickBorder,
  };

  // Format Season and Style header columns
  applyCellFormat(XLSX.utils.encode_cell({ r: 0, c: 0 }), emptyHeaderFormat);
  applyCellFormat(XLSX.utils.encode_cell({ r: 0, c: 1 }), emptyHeaderFormat);

  // Format month headers with merge
  for (let monthIdx = 0; monthIdx < monthLabels.length; monthIdx++) {
    const startCol = 2 + monthIdx * 3;
    const endCol = startCol + 2;
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: startCol });

    applyCellFormat(cellAddress, monthHeaderFormat);

    // Merge cells for this month
    if (!worksheet['!merges']) worksheet['!merges'] = [];
    worksheet['!merges'].push({
      s: { r: 0, c: startCol },
      e: { r: 0, c: endCol },
    });

    // Apply formatting to merged cells
    for (let c = startCol; c <= endCol; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      applyCellFormat(addr, {
        fill: { type: 'solid', fgColor: { rgb: 'FFD966' } },
        border: thickBorder,
      });
    }
  }

  // Format Total Sum header (now 3 columns, merged like every month block)
  applyCellFormat(XLSX.utils.encode_cell({ r: 0, c: totalPrevCol }), monthHeaderFormat);
  applyCellFormat(XLSX.utils.encode_cell({ r: 0, c: totalCurrCol }), monthHeaderFormat);
  applyCellFormat(XLSX.utils.encode_cell({ r: 0, c: totalVarCol }), monthHeaderFormat);

  if (!worksheet['!merges']) worksheet['!merges'] = [];
  worksheet['!merges'].push({
    s: { r: 0, c: totalPrevCol },
    e: { r: 0, c: totalVarCol },
  });

  // Format Row 2: Sub-headers (Season, Style, wk-33, wk-34, Variation, etc.)
  const subHeaderFormat = {
    fill: { type: 'solid', fgColor: { rgb: 'FFD966' } },
    font: { bold: true, size: 11 },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: thickBorder,
  };

  for (let col = 0; col < row2.length; col++) {
    applyCellFormat(XLSX.utils.encode_cell({ r: 1, c: col }), subHeaderFormat);
  }

  // Add Excel conditional formatting rules to variation columns (month
  // blocks + the Total Sum block).
  worksheet['!conditionalFormatting'] = [];
  const allVarCols = monthLabels.map((_, monthIdx) => 2 + monthIdx * 3 + 2);
  allVarCols.push(totalVarCol);
  for (const varCol of allVarCols) {
    const startRow = 3;
    const endRow = allRows.length;
    const varRange = `${XLSX.utils.encode_col(varCol)}${startRow}:${XLSX.utils.encode_col(varCol)}${endRow}`;

    worksheet['!conditionalFormatting'].push({
      ref: varRange,
      rules: [
        {
          type: 'cellIs',
          operator: 'lessThan',
          formula: '0',
          style: {
            fill: { patternType: 'solid', fgColor: { rgb: 'F4CCCC' } },
            font: { color: { rgb: 'C00000' } },
          },
        },
        {
          type: 'cellIs',
          operator: 'greaterThan',
          formula: '0',
          style: {
            fill: { patternType: 'solid', fgColor: { rgb: 'C6EFCE' } },
            font: { color: { rgb: '006100' } },
          },
        },
        {
          type: 'cellIs',
          operator: 'equal',
          formula: '0',
          style: {
            fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
            font: { color: { rgb: '000000' } },
          },
        },
      ],
    });
  }

  // Format data rows
  for (let rowIdx = 2; rowIdx < allRows.length; rowIdx++) {
    const row = allRows[rowIdx];
    const isSeasonTotal = row._isSeasonTotal;
    const isGrandTotal = row._isGrandTotal;

    for (let col = 0; col < row.length; col++) {
      if (col === 0 || col === 1) continue; // Skip Season and Style - will handle below

      const cellAddress = XLSX.utils.encode_cell({ r: rowIdx, c: col });
      let cellFormat = {
        alignment: { horizontal: 'right', vertical: 'center' },
        border: thinBorder,
        font: { size: 10 },
      };

      const colIndex = col - 2; // Adjust for Season/Style offset
      const monthIdx = Math.floor(colIndex / 3);
      const posInMonth = colIndex % 3; // 0=wk-33, 1=wk-34, 2=Variation

      const cellValue = row[col];

      // Apply formatting based on row type and column
      if (isSeasonTotal) {
        cellFormat.fill = { type: 'solid', fgColor: { rgb: 'D9D9D9' } }; // Light gray
        cellFormat.font = { bold: true, size: 10 };
        cellFormat.border = {
          ...thinBorder,
          top: { style: 'thick', color: { rgb: '000000' } },
        };
        cellFormat.alignment.horizontal = 'right';
      } else if (isGrandTotal) {
        cellFormat.fill = { type: 'solid', fgColor: { rgb: 'F4B183' } }; // Golden/dark yellow
        cellFormat.font = { bold: true, size: 11 };
        cellFormat.border = thickBorder;
        cellFormat.alignment.horizontal = 'right';
      } else if (monthIdx <= monthLabels.length && posInMonth === 2 && typeof cellValue === 'string' && cellValue.startsWith('=')) {
        cellFormat.font = { color: { rgb: '000000' }, size: 10 };
      }

      // Add number formatting for numeric values
      if (typeof cellValue === 'number' && !Number.isNaN(cellValue)) {
        cellFormat.numFmt = '#,##0';
      }

      applyCellFormat(cellAddress, cellFormat);
    }

    // Format Season and Style columns
    for (let col = 0; col <= 1; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIdx, c: col });
      let textFormat = {
        alignment: { horizontal: 'left', vertical: 'center' },
        border: thinBorder,
        font: { size: 10 },
      };

      if (isSeasonTotal) {
        textFormat.fill = { type: 'solid', fgColor: { rgb: 'D9D9D9' } };
        textFormat.font = { bold: true, size: 10 };
        textFormat.border = {
          ...thinBorder,
          top: { style: 'thick', color: { rgb: '000000' } },
        };
      } else if (isGrandTotal) {
        textFormat.fill = { type: 'solid', fgColor: { rgb: 'F4B183' } };
        textFormat.font = { bold: true, size: 11 };
        textFormat.border = thickBorder;
      }

      applyCellFormat(cellAddress, textFormat);
    }
  }

  return worksheet;
}

function generateNonDenimChartVariation(previousWeekBuffer, currentWeekBuffer, previousWeekFilename, currentWeekFilename, selection = {}) {
  // Parse files
  const previousRows = parseChartFile(previousWeekBuffer);
  const currentRows = parseChartFile(currentWeekBuffer);

  if (!previousRows.length) throw new Error('Previous week file has no data rows.');
  if (!currentRows.length) throw new Error('Current week file has no data rows.');

  // Extract week codes from filenames
  const previousWeekCode = extractWeekCodeFromFilename(previousWeekFilename);
  const currentWeekCode = extractWeekCodeFromFilename(currentWeekFilename);

  if (!previousWeekCode || !currentWeekCode) {
    throw new Error('Could not extract week codes from filenames. Expected format: WK-202633.xlsx or WK202633.xlsx');
  }

  // Aggregate data
  const { aggregated: previousAgg, months } = aggregateChartData(previousRows);
  const { aggregated: currentAgg, months: currentMonths } = aggregateChartData(currentRows);

  // Merge months from both weeks
  const allMonths = new Map();
  for (const m of months) allMonths.set(m.label, m);
  for (const m of currentMonths) allMonths.set(m.label, m);

  const sortedMonths = Array.from(allMonths.values()).sort((a, b) => {
    if (a.date.year !== b.date.year) return a.date.year - b.date.year;
    return a.date.month - b.date.month;
  });

  const selectedMonths = getMonthRange(sortedMonths, selection.startMonth, selection.endMonth);

  const previousSheet = generateChartSheet(previousWeekCode, previousAgg, selectedMonths, 'Previous Week');
  const currentSheet = generateChartSheet(currentWeekCode, currentAgg, selectedMonths, 'Current Week');
  const comparisonSheet = generateComparisonSheet(previousAgg, currentAgg, selectedMonths, previousWeekCode, currentWeekCode);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, previousSheet, previousWeekCode.toUpperCase());
  XLSX.utils.book_append_sheet(workbook, currentSheet, currentWeekCode.toUpperCase());
  XLSX.utils.book_append_sheet(workbook, comparisonSheet, 'Comparison');

  validateWorkbookMonthRange(workbook, selectedMonths.map((month) => month.label));

  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

module.exports = {
  generateNonDenimChartVariation,
  generateComparisonSheet,
  aggregateChartData,
  getMonthRange,
  extractWeekCodeFromFilename,
  parseChartFile,
};