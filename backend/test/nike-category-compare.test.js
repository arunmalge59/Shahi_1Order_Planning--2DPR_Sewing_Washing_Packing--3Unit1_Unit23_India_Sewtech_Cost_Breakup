const assert = require('node:assert/strict');
const { test } = require('node:test');
const ExcelJS = require('exceljs');
const {
  parseNikeSnapshot,
  buildComparison,
  generateCategoryComparisonExcel,
} = require('../modules/planning/src/nikeCategoryCompare');

const categories = ['BOTTOM', 'JACKET', 'TOP'];

function makeRows(monthValues) {
  return monthValues.flatMap(({ year, month, values }) =>
    categories.flatMap((category) => [
      {
        category,
        exFactory: new Date(Date.UTC(year, month, 5)),
        gac: new Date(Date.UTC(year, month - 1, 5)),
        qty: values[category].confirmed,
        status: 'Confirmed',
      },
      {
        category,
        exFactory: new Date(Date.UTC(year, month, 5)),
        gac: new Date(Date.UTC(year, month - 1, 5)),
        qty: values[category].projections,
        status: 'Projections',
      },
    ])
  );
}

async function makeSnapshotBuffer(
  rows,
  sheetName = 'Nike Data',
  headers = ['Category', 'EX FACTORY', 'Qty', 'Status', 'GAC', 'Style']
) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);
  worksheet.addRow(headers);
  rows.forEach((row) => {
    const values = {
      Category: row.category,
      'EX FACTORY': row.exFactory,
      Qty: row.qty,
      Status: row.status,
      GAC: row.gac,
      Style: row.gac instanceof Date ? `${row.category}-${row.gac.toISOString().slice(0, 7)}` : '',
    };
    worksheet.addRow(headers.map((header) => values[header]));
  });
  return workbook.xlsx.writeBuffer();
}

async function makeComparison(previousRows, currentRows) {
  const [previous, current] = await Promise.all([
    makeSnapshotBuffer(previousRows),
    makeSnapshotBuffer(currentRows),
  ]);
  const prevParsed = await parseNikeSnapshot(previous, 'previous.xlsx');
  const currParsed = await parseNikeSnapshot(current, 'current.xlsx');
  const keys = [...new Set([...prevParsed.rows, ...currParsed.rows].map((row) => row.monthKey))].sort();
  return {
    previous: prevParsed,
    current: currParsed,
    comparison: buildComparison(prevParsed, currParsed, keys[0], keys[keys.length - 1], '23-Sep', '14-Oct'),
  };
}

function cellFormula(cell) {
  assert.equal(typeof cell.value, 'object');
  return cell.value.formula;
}

test('category comparison labels every month with its source year and exports TOTAL formulas in all sections', async () => {
  const previousRows = makeRows([
    {
      year: 2026,
      month: 12,
      values: {
        BOTTOM: { confirmed: 100, projections: 25 },
        JACKET: { confirmed: 10, projections: 20 },
        TOP: { confirmed: 5, projections: 10 },
      },
    },
    {
      year: 2027,
      month: 1,
      values: {
        BOTTOM: { confirmed: 50, projections: 0 },
        JACKET: { confirmed: 0, projections: 30 },
        TOP: { confirmed: 4, projections: 6 },
      },
    },
  ]);
  const currentRows = makeRows([
    {
      year: 2026,
      month: 12,
      values: {
        BOTTOM: { confirmed: 120, projections: 30 },
        JACKET: { confirmed: 0, projections: 5 },
        TOP: { confirmed: 5, projections: 10 },
      },
    },
    {
      year: 2027,
      month: 1,
      values: {
        BOTTOM: { confirmed: 55, projections: 5 },
        JACKET: { confirmed: 0, projections: 0 },
        TOP: { confirmed: 4, projections: 6 },
      },
    },
  ]);

  const { previous, current, comparison } = await makeComparison(previousRows, currentRows);
  assert.ok(previous.warnings.some((warning) => warning.includes('Sheet "Summary sheet" not found')));
  assert.equal(previous.rows.length, previousRows.length);
  assert.equal(previous.rows[0].monthKey, '2026-12');
  assert.equal(previousRows[0].exFactory.getUTCMonth(), 0);
  assert.equal(comparison.sanityChecks.passed, true);
  assert.deepEqual(comparison.monthHeaders.map((month) => month.label), ["Dec'26", "Jan'27"]);

  const all = comparison.tables.find((table) => table.id === 'all');
  const confirmed = comparison.tables.find((table) => table.id === 'confirmed');
  const projections = comparison.tables.find((table) => table.id === 'projections');

  for (const category of categories) {
    const allRow = all.rows.find((row) => row.category === category);
    const confirmedRow = confirmed.rows.find((row) => row.category === category);
    const projectionsRow = projections.rows.find((row) => row.category === category);
    for (const month of comparison.monthHeaders) {
      assert.equal(allRow.months[month.key].prev, confirmedRow.months[month.key].prev + projectionsRow.months[month.key].prev);
      assert.equal(allRow.months[month.key].curr, confirmedRow.months[month.key].curr + projectionsRow.months[month.key].curr);
    }
  }

  const buffer = await generateCategoryComparisonExcel(comparison);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet('Category Comparison');

  const expectedTotals = {
    all: [[175, 210, 35], [60, 5, -55], [25, 25, 0], [260, 240, -20]],
    confirmed: [[150, 175, 25], [10, 0, -10], [9, 9, 0], [169, 184, 15]],
    projections: [[25, 35, 10], [50, 5, -45], [16, 16, 0], [91, 56, -35]],
  };

  for (const [tableIndex, tableId] of ['all', 'confirmed', 'projections'].entries()) {
    const titleRowNumber = Array.from({ length: worksheet.rowCount }, (_, index) => index + 1).find(
      (rowNumber) => String(worksheet.getCell(rowNumber, 1).value || '').startsWith(comparison.tables[tableIndex].title)
    );
    assert.ok(titleRowNumber, `${tableId} section exists`);
    const headerRow = titleRowNumber + 1;
    const firstDataRow = titleRowNumber + 3;
    const totalRow = firstDataRow + categories.length;

    assert.equal(worksheet.getCell(headerRow, 8).value, 'TOTAL');
    assert.equal(worksheet.getCell(headerRow, 2).value, "Dec'26");
    assert.equal(worksheet.getCell(headerRow, 5).value, "Jan'27");
    assert.equal(worksheet.getCell(headerRow + 1, 8).value, '23-Sep');
    assert.equal(worksheet.getCell(headerRow + 1, 9).value, '14-Oct');
    assert.equal(worksheet.getCell(headerRow + 1, 10).value, 'DIFF');

    for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex++) {
      const row = firstDataRow + categoryIndex;
      const diffCell = worksheet.getCell(row, 10);
      assert.deepEqual(
        [
          worksheet.getCell(row, 8).value.result ?? 0,
          worksheet.getCell(row, 9).value.result ?? 0,
          diffCell.value.result ?? 0,
        ],
        expectedTotals[tableId][categoryIndex]
      );
      assert.match(cellFormula(worksheet.getCell(row, 8)), /^SUM\(/);
      assert.match(cellFormula(worksheet.getCell(row, 9)), /^SUM\(/);
      assert.equal(cellFormula(worksheet.getCell(row, 10)), `I${row}-H${row}`);

      if (categories[categoryIndex] === 'BOTTOM') {
        assert.deepEqual(diffCell.fill.fgColor, { argb: 'FFC6EFCE' });
        assert.equal(diffCell.numFmt, '[>=10000000]##,##,##,##0;[>=100000]##,##,##0;##,##0');
      } else if (categories[categoryIndex] === 'JACKET') {
        assert.deepEqual(diffCell.fill.fgColor, { argb: 'FFFFC7CE' });
        assert.equal(diffCell.numFmt, '[>=10000000]##,##,##,##0;[>=100000]##,##,##0;##,##0');
      } else {
        assert.equal(diffCell.numFmt, '"-"');
        assert.deepEqual(diffCell.fill, { type: 'pattern', pattern: 'none' });
      }
    }

    const totalDiffCell = worksheet.getCell(totalRow, 10);
    assert.deepEqual(
      [
        worksheet.getCell(totalRow, 8).value.result ?? 0,
        worksheet.getCell(totalRow, 9).value.result ?? 0,
        totalDiffCell.value.result ?? 0,
      ],
      expectedTotals[tableId][3]
    );
    assert.equal(cellFormula(worksheet.getCell(totalRow, 8)), `SUM(H${firstDataRow}:H${totalRow - 1})`);
    assert.equal(cellFormula(worksheet.getCell(totalRow, 9)), `SUM(I${firstDataRow}:I${totalRow - 1})`);
    assert.equal(cellFormula(worksheet.getCell(totalRow, 10)), `I${totalRow}-H${totalRow}`);
    assert.equal(
      totalDiffCell.numFmt,
      expectedTotals[tableId][3][2] === 0
        ? '"-"'
        : '[>=10000000]##,##,##,##0;[>=100000]##,##,##0;##,##0'
    );
    if (expectedTotals[tableId][3][2] > 0) {
      assert.deepEqual(totalDiffCell.fill.fgColor, { argb: 'FFC6EFCE' });
    } else if (expectedTotals[tableId][3][2] < 0) {
      assert.deepEqual(totalDiffCell.fill.fgColor, { argb: 'FFFFC7CE' });
    } else {
      assert.deepEqual(totalDiffCell.fill, { type: 'pattern', pattern: 'none' });
    }
  }

  assert.equal(cellFormula(worksheet.getCell(4, 4)), 'C4-B4');
  assert.equal(worksheet.getCell(4, 4).value.result, 25);
  assert.equal(cellFormula(worksheet.getCell(4, 7)), 'F4-E4');
  assert.equal(worksheet.getCell(4, 7).value.result, 10);
});

test('sixteen monthly headings place the TOTAL group in columns AX:AZ', async () => {
  const months = [];
  for (let i = 0; i < 16; i++) {
    const date = new Date(Date.UTC(2026, 3 + i, 5));
    months.push({
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      values: {
        BOTTOM: { confirmed: 10, projections: 5 },
        JACKET: { confirmed: 20, projections: 10 },
        TOP: { confirmed: 30, projections: 15 },
      },
    });
  }

  const { comparison } = await makeComparison(makeRows(months), makeRows(months));
  assert.deepEqual(
    comparison.monthHeaders.map((month) => month.label),
    ["Apr'26", "May'26", "Jun'26", "Jul'26", "Aug'26", "Sep'26", "Oct'26", "Nov'26",
      "Dec'26", "Jan'27", "Feb'27", "Mar'27", "Apr'27", "May'27", "Jun'27", "Jul'27"]
  );

  const buffer = await generateCategoryComparisonExcel(comparison);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet('Category Comparison');
  for (const table of comparison.tables) {
    const titleRow = Array.from({ length: worksheet.rowCount }, (_, index) => index + 1).find(
      (rowNumber) => String(worksheet.getCell(rowNumber, 1).value || '').startsWith(table.title)
    );
    assert.ok(titleRow);
    const headerRow = titleRow + 1;
    const firstDataRow = titleRow + 3;
    const totalRow = firstDataRow + categories.length;
    assert.equal(worksheet.getCell(headerRow, 50).value, 'TOTAL');
    assert.equal(worksheet.getCell(headerRow, 50).address, `AX${headerRow}`);
    assert.equal(worksheet.getCell(headerRow, 51).address, `AY${headerRow}`);
    assert.equal(worksheet.getCell(headerRow, 52).address, `AZ${headerRow}`);

    for (let row = firstDataRow; row <= totalRow; row++) {
      assert.equal(cellFormula(worksheet.getCell(row, 52)), `AY${row}-AX${row}`);
    }
  }
});

test('unclassified statuses remain in ALL and are reported as an ALL breakdown discrepancy', async () => {
  const rows = makeRows([
    {
      year: 2026,
      month: 4,
      values: {
        BOTTOM: { confirmed: 10, projections: 5 },
        JACKET: { confirmed: 20, projections: 10 },
        TOP: { confirmed: 30, projections: 15 },
      },
    },
  ]);
  rows[0].status = 'Pending';

  const { previous, comparison } = await makeComparison(rows, rows);
  assert.ok(previous.warnings.some((warning) => warning.includes('unclassified status "pending"')));
  assert.equal(comparison.tables.find((table) => table.id === 'all').rows[0].months['2026-04'].prev, 15);
  assert.equal(comparison.sanityChecks.passed, false);
  assert.ok(comparison.sanityChecks.errors.some((error) => error.includes('Confirmed') && error.includes('BOTTOM - 2026-04')));
});

test('GAC is required, EX FACTORY is optional, and GAC dates determine the month', async () => {
  const buffer = await makeSnapshotBuffer([
    { category: 'BOTTOM', gac: '2026-10-14', qty: 100 },
    { category: 'JACKET', gac: '2026-11-02', qty: 50 },
  ], 'GAC data', ['Category', 'GAC', 'Qty']);

  const parsed = await parseNikeSnapshot(buffer, 'GAC-only.xlsx');
  assert.deepEqual(parsed.rows.map((row) => row.monthKey), ['2026-10', '2026-11']);
  assert.deepEqual(parsed.rows.map((row) => row.qty), [100, 50]);
  assert.ok(parsed.warnings.some((warning) => warning.includes('missing a "Status" column')));
  assert.ok(!parsed.warnings.some((warning) => warning.includes('EX FACTORY')));

  const missingGacHeader = await makeSnapshotBuffer([
    { category: 'BOTTOM', exFactory: new Date(Date.UTC(2026, 9, 14)), qty: 100 },
  ], 'Missing GAC', ['Category', 'EX FACTORY', 'Qty']);
  await assert.rejects(
    parseNikeSnapshot(missingGacHeader, 'missing-gac.xlsx'),
    /Required columns are Category, GAC, and Qty/
  );
});

test('missing and invalid GAC row dates are clearly warned and excluded from month totals', async () => {
  const buffer = await makeSnapshotBuffer([
    { category: 'BOTTOM', gac: null, qty: 40 },
    { category: 'JACKET', gac: '2026-02-31', qty: 60 },
    { category: 'TOP', gac: '2026-10-14', qty: 20 },
  ], 'GAC audit', ['Category', 'GAC', 'Qty']);

  const parsed = await parseNikeSnapshot(buffer, 'gac-audit.xlsx');
  assert.deepEqual(parsed.rows.map((row) => row.monthKey), ['2026-10']);
  assert.equal(parsed.rawTotalQty, 120);
  assert.equal(parsed.unassignedQty, 100);
  assert.ok(parsed.warnings.some((warning) => warning.includes('a missing GAC date')), JSON.stringify(parsed.warnings));
  assert.ok(parsed.warnings.some((warning) => warning.includes('an invalid GAC date')), JSON.stringify(parsed.warnings));
  assert.ok(parsed.warnings.some((warning) => warning.includes('100 qty') && warning.includes('no valid GAC date')));
});
