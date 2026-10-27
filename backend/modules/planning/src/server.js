require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');

const { parseSnapshotFile } = require('./csvHelper');
const { compareSnapshots } = require('./compare');
const { generateOrderFile } = require('./orderFileGenerator');
const { generateHmOrderFile } = require('./hmOrderFileGenerator');
const { generateHmNonDenimCpFile, generateHmNonDenimOrderFile } = require('./hmNonDenimOrderFileGenerator');
const { generateNonDenimChartVariation, aggregateChartData } = require('./hmNonDenimChartVariation');
const { generateAdidasOrderFile } = require('./adidasOrderFileGenerator');
const {
  parseNikeSnapshot,
  computeAvailableMonths,
  buildComparison,
  generateCategoryComparisonExcel,
} = require('./nikeCategoryCompare');
const { pool, initDb } = require('./db');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
app.use(cors());
app.use(express.json());

function getUploadedFile(req, fieldName) {
  if (Array.isArray(req.files)) {
    return req.files.find((f) => f.fieldname === fieldName);
  }
  if (req.files && req.files[fieldName]) {
    return req.files[fieldName][0];
  }
  return req.file || null;
}

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Main endpoint: accepts two CSV files (multipart/form-data),
// fields must be named "julyFile" (older snapshot) and
// "augustFile" (newer snapshot).
app.post(
  '/api/compare',
  upload.any(),
  async (req, res) => {
    try {
      const julyFile = getUploadedFile(req, 'julyFile');
      const augustFile = getUploadedFile(req, 'augustFile');

      if (!julyFile || !augustFile) {
        return res.status(400).json({
          error: 'Both julyFile (older snapshot) and augustFile (newer snapshot) are required.',
        });
      }

      const julyRows = parseSnapshotFile(julyFile.buffer, julyFile.originalname);
      const augustRows = parseSnapshotFile(augustFile.buffer, augustFile.originalname);

      const result = compareSnapshots(julyRows, augustRows);

      // Save a record of this comparison run. If the database is not
      // reachable, the comparison still succeeds ??? the run just won't
      // be saved to history.
      try {
        const runRes = await pool.query(
          `INSERT INTO comparison_runs
            (older_file_name, newer_file_name, total_styles, changed_styles, qty_changed, pcd_changed, ex_factory_changed)
           VALUES (?,?,?,?,?,?,?)`,
          [
            julyFile.originalname,
            augustFile.originalname,
            result.summary.totalStyles,
            result.summary.changedStyles,
            result.summary.qtyChanged,
            result.summary.pcdChanged,
            result.summary.exFactoryChanged,
          ]
        );
        const runId = runRes[0].insertId;

        for (const s of result.styles) {
          await pool.query(
            `INSERT INTO comparison_style_changes
              (run_id, style, july_qty, august_qty, qty_delta, pcd_changed, ex_factory_changed, status)
             VALUES (?,?,?,?,?,?,?,?)`,
            [runId, s.style, s.julyQty, s.augustQty, s.qtyDelta, s.pcdChanged, s.exFactoryChanged, s.status]
          );
        }

        result.runId = runId;
      } catch (dbErr) {
        console.error('Could not save comparison run to database:', dbErr.message);
      }

      res.json(result);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: 'Failed to compare files.', details: err.message });
    }
  }
);

app.post('/api/order-file/generate', upload.any(), (req, res) => {
  try {
    const file = getUploadedFile(req, 'file') || req.files?.[0];
    if (!file) {
      return res.status(400).json({ error: 'An Excel workbook is required in the file field.' });
    }
    const csv = generateOrderFile(file.buffer, req.body.weekCode || 'W36');
    res.set({
      'Content-Type': 'text/csv; charset=utf-8',
      'Content-Disposition': 'attachment; filename="nike-order-file.csv"',
    });
    res.send(csv);
  } catch (err) {
    console.error('Could not generate Nike order file:', err.message);
    res.status(400).json({ error: err.message || 'Failed to generate the order file.' });
  }
});

// UPDATED: now accepts an optional second field "frFile" (the FR style/
// description workbook) alongside the required "file" (H&M planning workbook).
// When frFile is provided, generateHmOrderFile matches Season+Style against it
// to fill O.PROD / O.UDITEM / O.Descrip. When it's omitted, behavior is exactly
// what it was before (raw style text, no FR matching, nothing skipped).
app.post(
  '/api/hm/order-file',
  upload.any(),
  (req, res) => {
    try {
      const file = getUploadedFile(req, 'file') || req.files?.[0];
      const frFile = getUploadedFile(req, 'frFile');

      if (!file) {
        return res.status(400).json({ error: 'An H&M Excel workbook is required in the file field.' });
      }
      const result = generateHmOrderFile(file.buffer, req.body.weekCode, frFile?.buffer);
      res.json(result);
    } catch (err) {
      console.error('Could not generate H&M order file:', err.message);
      res.status(400).json({ error: err.message || 'Failed to generate the H&M order file.' });
    }
  }
);

app.post('/api/hm/order-file-non-denim/cp-template', upload.any(), async (req, res) => {
  try {
    const marketingFile = getUploadedFile(req, 'marketingFile') || getUploadedFile(req, 'file');
    const previousCpFile = getUploadedFile(req, 'previousCpFile') || getUploadedFile(req, 'previousWeekFile');

    if (!marketingFile) {
      return res.status(400).json({ error: 'A Marketing Excel workbook is required in the marketingFile field.' });
    }

    const result = await generateHmNonDenimCpFile(marketingFile.buffer, previousCpFile?.buffer || null);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="hm-non-denim-cp-file.xlsx"',
      'X-Generated-CP-Rows': String(result.summary.generatedCpRows),
      'X-Auto-Filled': String(result.summary.autoFilled || 0),
      'X-Flagged-Item-Known': String(result.summary.flaggedItemKnown || 0),
      'X-Flagged-New-Item': String(result.summary.flaggedNewItem || 0),
      'X-Flagged-For-Review': String(result.summary.flaggedForManualReview || 0),
    });

    // If there are flagged rows, include a summary in response headers
    if (result.flaggedRows && result.flaggedRows.length > 0) {
      const flaggedSummary = result.flaggedRows.slice(0, 5).map((f) =>
        `ItemNo:${f.itemNo} (${f.flagType === 'orange' ? 'Item Known' : 'Brand New'})`
      ).join(' | ');
      res.set('X-Flagged-Summary', flaggedSummary + (result.flaggedRows.length > 5 ? '...' : ''));
    }

    res.send(result.buffer);
  } catch (err) {
    console.error('Could not generate H&M Non-Denim CP file:', err.message);
    res.status(400).json({ error: err.message || 'Failed to generate the H&M Non-Denim CP file.' });
  }
});
app.post('/api/hm/order-file-non-denim', upload.any(), (req, res) => {
  try {
    const marketingFile = getUploadedFile(req, 'marketingFile') || getUploadedFile(req, 'file');
    const cpFile = getUploadedFile(req, 'cpFile');
    const previousCpFile = getUploadedFile(req, 'previousCpFile') || getUploadedFile(req, 'previousWeekFile');

    if (!marketingFile || !cpFile) {
      return res.status(400).json({
        error: 'Both marketingFile and cpFile are required in the request.',
      });
    }

    const result = generateHmNonDenimOrderFile(
      marketingFile.buffer,
      cpFile.buffer,
      req.body.weekCode,
      previousCpFile?.buffer || null
    );
    res.json(result);
  } catch (err) {
    console.error('Could not generate H&M Non-Denim order file:', err.message);
    res.status(400).json({ error: err.message || 'Failed to generate the H&M Non-Denim order file.' });
  }
});

app.post('/api/hm/non-denim/chart-variation/months', upload.any(), (req, res) => {
  try {
    const previousWeekFile = getUploadedFile(req, 'previousWeekFile') || getUploadedFile(req, 'previousCpFile');
    const currentWeekFile = getUploadedFile(req, 'currentWeekFile') || getUploadedFile(req, 'currentCpFile');

    if (!previousWeekFile || !currentWeekFile) {
      return res.status(400).json({
        error: 'Both previousWeekFile and currentWeekFile are required in the request.',
      });
    }

    const previousRows = require('./hmNonDenimChartVariation').parseChartFile(previousWeekFile.buffer);
    const currentRows = require('./hmNonDenimChartVariation').parseChartFile(currentWeekFile.buffer);

    const previousAgg = aggregateChartData(previousRows).months;
    const currentAgg = aggregateChartData(currentRows).months;

    const allMonths = new Map();
    for (const month of previousAgg) allMonths.set(month.label, month);
    for (const month of currentAgg) allMonths.set(month.label, month);

    const months = Array.from(allMonths.values()).sort((a, b) => {
      if (a.date.year !== b.date.year) return a.date.year - b.date.year;
      return a.date.month - b.date.month;
    });

    res.json({ months: months.map((month) => month.label) });
  } catch (err) {
    console.error('Could not derive month list:', err.message);
    res.status(400).json({ error: err.message || 'Failed to derive available months.' });
  }
});

app.post('/api/hm/non-denim/chart-variation', upload.any(), (req, res) => {
  try {
    const previousWeekFile = getUploadedFile(req, 'previousWeekFile') || getUploadedFile(req, 'previousCpFile');
    const currentWeekFile = getUploadedFile(req, 'currentWeekFile') || getUploadedFile(req, 'currentCpFile');

    if (!previousWeekFile || !currentWeekFile) {
      return res.status(400).json({
        error: 'Both previousWeekFile and currentWeekFile are required in the request.',
      });
    }

    const selection = {
      startMonth: req.body.startMonth || '',
      endMonth: req.body.endMonth || '',
    };

    const fileBuffer = generateNonDenimChartVariation(
      previousWeekFile.buffer,
      currentWeekFile.buffer,
      previousWeekFile.originalname,
      currentWeekFile.originalname,
      selection,
    );

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="H&M_Chart_Variation_Non-Denim.xlsx"',
    });
    res.send(fileBuffer);
  } catch (err) {
    console.error('Could not generate H&M Non-Denim chart variation:', err.message);
    res.status(400).json({ error: err.message || 'Failed to generate the chart variation file.' });
  }
});

app.post('/api/adidas/order-file', upload.any(), (req, res) => {
  try {
    const file = getUploadedFile(req, 'file') || req.files?.[0];
    if (!file) {
      return res.status(400).json({ error: 'An Order Chart Excel workbook is required in the file field.' });
    }
    const result = generateAdidasOrderFile(file.buffer, req.body.weekCode || 'WK-33 Rev1');
    res.json(result);
  } catch (err) {
    console.error('Could not generate Adidas order file:', err.message);
    res.status(400).json({ error: err.message || 'Failed to generate the Adidas order file.' });
  }
});

app.post('/api/nike-category-compare/months', upload.any(), async (req, res) => {
  try {
    const previousFile = getUploadedFile(req, 'previousFile') || getUploadedFile(req, 'prevFile') || getUploadedFile(req, 'julyFile');
    const currentFile = getUploadedFile(req, 'currentFile') || getUploadedFile(req, 'currFile') || getUploadedFile(req, 'augustFile');

    if (!previousFile || !currentFile) {
      return res.status(400).json({ error: 'Both Previous Snapshot and Current Snapshot files are required.' });
    }

    const prevParsed = await parseNikeSnapshot(previousFile.buffer, previousFile.originalname);
    const currParsed = await parseNikeSnapshot(currentFile.buffer, currentFile.originalname);

    const monthData = computeAvailableMonths(prevParsed, currParsed);
    const isSameFile = previousFile.originalname === currentFile.originalname && previousFile.size === currentFile.size;

    res.json({
      ...monthData,
      isSameFile,
      warnings: [...(prevParsed.warnings || []), ...(currParsed.warnings || [])],
    });
  } catch (err) {
    console.error('Could not derive Nike Category months:', err.message);
    res.status(400).json({ error: err.message || 'Failed to parse files for available months.' });
  }
});

app.post('/api/nike-category-compare/preview', upload.any(), async (req, res) => {
  try {
    const previousFile = getUploadedFile(req, 'previousFile') || getUploadedFile(req, 'prevFile') || getUploadedFile(req, 'julyFile');
    const currentFile = getUploadedFile(req, 'currentFile') || getUploadedFile(req, 'currFile') || getUploadedFile(req, 'augustFile');

    if (!previousFile || !currentFile) {
      return res.status(400).json({ error: 'Both Previous Snapshot and Current Snapshot files are required.' });
    }

    const prevParsed = await parseNikeSnapshot(previousFile.buffer, previousFile.originalname);
    const currParsed = await parseNikeSnapshot(currentFile.buffer, currentFile.originalname);

    const prevLabel = req.body.prevLabel || 'Previous';
    const currLabel = req.body.currLabel || 'Current';
    const fromMonth = req.body.fromMonth || '';
    const toMonth = req.body.toMonth || '';

    const comparison = buildComparison(prevParsed, currParsed, fromMonth, toMonth, prevLabel, currLabel);
    const isSameFile = previousFile.originalname === currentFile.originalname && previousFile.size === currentFile.size;

    res.json({
      ...comparison,
      isSameFile,
    });
  } catch (err) {
    console.error('Could not generate Nike Category preview:', err.message);
    res.status(400).json({ error: err.message || 'Failed to compare snapshots.' });
  }
});

app.post('/api/nike-category-compare/export', upload.any(), async (req, res) => {
  try {
    const previousFile = getUploadedFile(req, 'previousFile') || getUploadedFile(req, 'prevFile') || getUploadedFile(req, 'julyFile');
    const currentFile = getUploadedFile(req, 'currentFile') || getUploadedFile(req, 'currFile') || getUploadedFile(req, 'augustFile');

    if (!previousFile || !currentFile) {
      return res.status(400).json({ error: 'Both Previous Snapshot and Current Snapshot files are required.' });
    }

    const prevParsed = await parseNikeSnapshot(previousFile.buffer, previousFile.originalname);
    const currParsed = await parseNikeSnapshot(currentFile.buffer, currentFile.originalname);

    const prevLabel = req.body.prevLabel || 'Previous';
    const currLabel = req.body.currLabel || 'Current';
    const fromMonth = req.body.fromMonth || '';
    const toMonth = req.body.toMonth || '';

    const comparison = buildComparison(prevParsed, currParsed, fromMonth, toMonth, prevLabel, currLabel);
    const fileBuffer = await generateCategoryComparisonExcel(comparison);

    const safePrevLabel = String(prevLabel).replace(/[^a-zA-Z0-9_-]/g, '_');
    const safeCurrLabel = String(currLabel).replace(/[^a-zA-Z0-9_-]/g, '_');
    const exportFileName = `Nike_Category_Comparison_${safePrevLabel}_vs_${safeCurrLabel}.xlsx`;

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${exportFileName}"`,
    });
    res.send(fileBuffer);
  } catch (err) {
    console.error('Could not export Nike Category Comparison Excel:', err.message);
    res.status(400).json({ error: err.message || 'Failed to generate category comparison Excel.' });
  }
});

// Returns past comparison runs (history).
app.get('/api/runs', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT * FROM comparison_runs ORDER BY created_at DESC LIMIT 20`
    );
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: 'Failed to fetch run history.', details: err.message });
  }
});

module.exports = app;