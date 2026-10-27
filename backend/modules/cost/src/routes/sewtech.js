const express = require('express');
const multer = require('multer');
const pool = require('../db/pool');
const engine = require('../services/sewtechEngine');
const excelIO = require('../services/sewtechExcelIO');

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 35 * 1024 * 1024 } });

// Helper: load buyer map from DB
async function loadBuyerMap() {
  const { rows } = await pool.query(`
    SELECT m.id, m.raw_buyer_norm, m.raw_buyer_display, g.name AS group_name, g.id AS buyer_group_id
    FROM sewtech_buyer_map m
    JOIN sewtech_buyer_groups g ON m.buyer_group_id = g.id
  `);
  const map = new Map();
  for (const r of rows) {
    map.set(r.raw_buyer_norm, { group_name: r.group_name, buyer_group_id: r.buyer_group_id, id: r.id });
  }
  return map;
}

// Helper: load unit map from DB
async function loadUnitMap() {
  const { rows } = await pool.query('SELECT title_number, unit_code FROM sewtech_unit_map');
  const map = new Map();
  for (const r of rows) {
    map.set(Number(r.title_number), r.unit_code);
  }
  return map;
}

// POST /api/sewtech/upload-raw (multipart "file")
// Step 1: Save file, scan tabs, return detected dates and ignored tabs. No rows created yet.
router.post('/upload-raw', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded (expected field "file").' });
    }

    const { validTabs, ignoredTabs, duplicateDates } = engine.scanWorkbookTabs(req.file.buffer);

    if (validTabs.length === 0) {
      return res.status(400).json({
        error: 'No valid daily production tabs found in the uploaded workbook.',
        ignoredTabs,
      });
    }

    // Create preliminary batch with raw file saved
    const minDate = validTabs[0].date;
    const maxDate = validTabs[validTabs.length - 1].date;
    const defaultLabel = excelIO.formatOrdinalDateRange(minDate, maxDate);

    const insertRes = await pool.query(
      `INSERT INTO sewtech_batches (
         week_label, start_date, end_date, working_days, working_days_source,
         raw_filename, raw_file, status
       ) VALUES (?, ?, ?, ?, 'auto', ?, ?, 'raw_uploaded')
       `,
      [defaultLabel, minDate, maxDate, validTabs.length, req.file.originalname, req.file.buffer]
    );

    const batchId = insertRes.insertId;

    res.json({
      batch_id: batchId,
      raw_filename: req.file.originalname,
      min_date: minDate,
      max_date: maxDate,
      valid_tabs: validTabs.map((t) => ({ sheet_name: t.sheetName, date: t.date })),
      ignored_tabs: ignoredTabs,
      duplicate_dates: duplicateDates,
    });
  } catch (err) {
    console.error('Sewtech upload-raw error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/sewtech/:id/process  { start_date, end_date, working_days? }
// Step 2: Parse tabs in range, compute metrics, apply buyer mapping, check pending items
router.post('/:id/process', async (req, res) => {
  try {
    const { id } = req.params;
    const { start_date, end_date, working_days } = req.body;

    if (!start_date || !end_date) {
      return res.status(400).json({ error: 'start_date and end_date are required (YYYY-MM-DD).' });
    }

    const batchRes = await pool.query('SELECT * FROM sewtech_batches WHERE id=?', [id]);
    const batch = batchRes.rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });
    if (!batch.raw_file) return res.status(400).json({ error: 'No raw file found for this batch.' });

    const [buyerMap, unitMap] = await Promise.all([loadBuyerMap(), loadUnitMap()]);

    const result = engine.processWorkbook(batch.raw_file, {
      startDate: start_date,
      endDate: end_date,
      workingDays: working_days ? Number(working_days) : null,
      includeHp: batch.include_hp,
      buyerMap,
      unitMap,
    });

    const weekLabel = excelIO.formatOrdinalDateRange(start_date, end_date);
    const workingDaysSource = working_days ? 'user_confirmed' : 'auto';

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      // Update batch
      await client.query(
        `UPDATE sewtech_batches
         SET week_label = ?, start_date = ?, end_date = ?,
             working_days = ?, working_days_source = ?, status = 'processed'
         WHERE id = ?`,
        [weekLabel, start_date, end_date, result.finalWorkingDays, workingDaysSource, id]
      );

      // Clean existing rows & logs
      await client.query('DELETE FROM sewtech_rows WHERE batch_id=?', [id]);
      await client.query('DELETE FROM sewtech_change_log WHERE batch_id=?', [id]);
      await client.query('DELETE FROM sewtech_final_summary WHERE batch_id=?', [id]);

      // Insert rows
      for (const row of result.rows) {
        await client.query(
          `INSERT INTO sewtech_rows (
             batch_id, work_date, line, unit, operators_op, helpers_hp, operators_final,
             reg_hours, ot_man, ot_hr, raw_buyer, buyer_group, is_pending, pending_reason,
             style, item_no, order_qty, smv, earn_minutes, production, effi, mc_hrs,
             source_tab, source_row_no
           ) VALUES (
             ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
           )`,
          [
            id, row.work_date, row.line, row.unit, row.operators_op, row.helpers_hp, row.operators_final,
            row.reg_hours, row.ot_man, row.ot_hr, row.raw_buyer, row.buyer_group, row.is_pending, row.pending_reason,
            row.style, row.item_no, row.order_qty, row.smv, row.earn_minutes, row.production, row.effi, row.mc_hrs,
            row.source_tab, row.source_row_no,
          ]
        );
      }

      // Insert change log
      for (const entry of result.changeLogs) {
        await client.query(
          `INSERT INTO sewtech_change_log (batch_id, category, details) VALUES (?, ?, ?)`,
          [id, entry.category, entry.details || {}]
        );
      }

      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    res.json({
      batch_id: id,
      week_label: weekLabel,
      start_date,
      end_date,
      working_days: result.finalWorkingDays,
      found_dates: result.foundDates,
      missing_dates: result.missingDates,
      row_count: result.rows.length,
      unknown_buyers: result.unknownBuyers,
      total_hp_count: result.totalHpCount,
      pending_hp_decision: result.pendingHpDecision,
      total_zero_prod_rows: result.totalZeroProdRows,
      total_zero_prod_operators: result.totalZeroProdOperators,
      validations: result.validations,
    });
  } catch (err) {
    console.error('Sewtech process error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/sewtech/:id/report
router.get('/:id/report', async (req, res) => {
  try {
    const { id } = req.params;
    const batchRes = await pool.query('SELECT * FROM sewtech_batches WHERE id=?', [id]);
    const batch = batchRes.rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    // Aggregate totals
    const totalsRes = await pool.query(
      `SELECT
         COUNT(*) AS total_rows,
         COALESCE(SUM(production), 0) AS total_production,
         COALESCE(SUM(operators_final), 0) AS total_operators,
         COUNT(IF(production = 0, 1, NULL)) AS zero_prod_rows,
         COALESCE(SUM(IF(production = 0, operators_final, 0)), 0) AS zero_prod_operators,
         COUNT(IF(is_pending = TRUE, 1, NULL)) AS pending_rows,
         COUNT(IF(helpers_hp > 0, 1, NULL)) AS hp_rows
       FROM sewtech_rows WHERE batch_id=?`,
      [id]
    );
    const totals = totalsRes.rows[0];

    // Day-by-day totals
    const dailyRes = await pool.query(
      `SELECT
         work_date,
         COUNT(*) AS rows_count,
         COALESCE(SUM(production), 0) AS day_production,
         COALESCE(SUM(operators_final), 0) AS day_operators
       FROM sewtech_rows
       WHERE batch_id=?
       GROUP BY work_date
       ORDER BY work_date`,
      [id]
    );

    // Pending buyers
    const pendingBuyersRes = await pool.query(
      `SELECT buyer_counts.raw_buyer, buyer_counts.count,
              JSON_ARRAYAGG(DATE_FORMAT(pending_dates.work_date, '%Y-%m-%d')) AS dates
       FROM (
         SELECT raw_buyer, COUNT(*) AS count
         FROM sewtech_rows
         WHERE batch_id=? AND is_pending=TRUE
         GROUP BY raw_buyer
       ) AS buyer_counts
       JOIN (
         SELECT DISTINCT raw_buyer, work_date
         FROM sewtech_rows
         WHERE batch_id=? AND is_pending=TRUE
       ) AS pending_dates ON pending_dates.raw_buyer <=> buyer_counts.raw_buyer
       GROUP BY buyer_counts.raw_buyer, buyer_counts.count`,
      [id, id]
    );

    const pendingBuyers = pendingBuyersRes.rows;

    // Logs
    const logRes = await pool.query(
      `SELECT category, details, created_at FROM sewtech_change_log WHERE batch_id=? ORDER BY id`,
      [id]
    );

    // Final summary if finalized
    let finalSummary = null;
    if (batch.status === 'finalized') {
      const summaryRes = await pool.query(
        `SELECT buyer_group, sum_production, sum_operators, avg_prod_per_day, avg_optrs_per_day
         FROM sewtech_final_summary
         WHERE batch_id=?
         ORDER BY buyer_group`,
        [id]
      );
      finalSummary = summaryRes.rows;
    }

    res.json({
      batch,
      totals,
      daily_totals: dailyRes.rows,
      pending_buyers: pendingBuyers,
      pending_hp_decision: Number(totals.hp_rows) > 0 && batch.include_hp == null,
      logs: logRes.rows,
      final_summary: finalSummary,
    });
  } catch (err) {
    console.error('Sewtech report error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/sewtech/:id/rows
router.get('/:id/rows', async (req, res) => {
  try {
    const { id } = req.params;
    const { limit = 300, offset = 0 } = req.query;

    const { rows } = await pool.query(
      `SELECT * FROM sewtech_rows
       WHERE batch_id=?
       ORDER BY work_date ASC, unit ASC, line ASC
       LIMIT ? OFFSET ?`,
      [id, Number(limit), Number(offset)]
    );

    const countRes = await pool.query('SELECT COUNT(*) AS count FROM sewtech_rows WHERE batch_id=?', [id]);

    res.json({
      total: Number(countRes.rows[0].count),
      rows,
    });
  } catch (err) {
    console.error('Sewtech rows error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/sewtech/:id/export-merged
router.get('/:id/export-merged', async (req, res) => {
  try {
    const { id } = req.params;
    const batchRes = await pool.query('SELECT * FROM sewtech_batches WHERE id=?', [id]);
    const batch = batchRes.rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    const { rows } = await pool.query(
      `SELECT * FROM sewtech_rows
       WHERE batch_id=?
       ORDER BY work_date ASC, unit ASC, line ASC`,
      [id]
    );

    const mappedRows = rows.map((r) => ({
      ...r,
      has_hp_pending: Number(r.helpers_hp) > 0 && batch.include_hp == null,
    }));

    const buffer = await excelIO.writeMergedSheet(mappedRows, { weekLabel: batch.week_label });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Sewtech_Merged_Batch_${id}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error('Sewtech export-merged error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/sewtech/:id/assign-buyer  { raw_buyer, buyer_group | new_group, save_mapping }
router.post('/:id/assign-buyer', async (req, res) => {
  try {
    const { id } = req.params;
    const { raw_buyer, buyer_group, new_group, save_mapping } = req.body;

    const targetGroup = engine.normalizeText(new_group || buyer_group);
    const rawDisplay = engine.normalizeText(raw_buyer);
    const rawNorm = engine.normalizeKey(raw_buyer);

    if (!rawDisplay || !targetGroup) {
      return res.status(400).json({ error: 'raw_buyer and buyer_group (or new_group) are required.' });
    }

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      // Ensure buyer group exists
      let groupRes = await client.query('SELECT id FROM sewtech_buyer_groups WHERE name=?', [targetGroup]);
      let groupId;
      if (groupRes.rows.length === 0) {
        await client.query('INSERT INTO sewtech_buyer_groups (name) VALUES (?)', [targetGroup]);
        const createGroup = await client.query(
          'SELECT id FROM sewtech_buyer_groups WHERE name=?',
          [targetGroup]
        );
        groupId = createGroup.rows[0].id;
      } else {
        groupId = groupRes.rows[0].id;
      }

      // Save to mapping master if requested
      if (save_mapping !== false) {
        await client.query(
          `INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id, updated_at)
           VALUES (?, ?, ?, CURRENT_TIMESTAMP)
           ON DUPLICATE KEY UPDATE buyer_group_id=VALUES(buyer_group_id), updated_at=CURRENT_TIMESTAMP`,
          [rawNorm, rawDisplay, groupId]
        );
      }

      // Update all matching rows in this batch
      await client.query(
        `UPDATE sewtech_rows
         SET buyer_group=?, is_pending=FALSE, pending_reason=NULL
         WHERE batch_id=? AND LOWER(TRIM(REGEXP_REPLACE(raw_buyer, '[[:space:]]+', ' '))) = ?`,
        [targetGroup, id, rawNorm]
      );

      // Change log entry
      await client.query(
        `INSERT INTO sewtech_change_log (batch_id, category, details)
         VALUES (?, 'buyer_mapped', ?)`,
        [id, { raw_buyer: rawDisplay, buyer_group: targetGroup, saved_to_master: save_mapping !== false }]
      );

      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    res.json({ ok: true, buyer_group: targetGroup });
  } catch (err) {
    console.error('Sewtech assign-buyer error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/sewtech/:id/hp-decision  { include_hp: boolean }
router.post('/:id/hp-decision', async (req, res) => {
  try {
    const { id } = req.params;
    const { include_hp } = req.body;

    if (typeof include_hp !== 'boolean') {
      return res.status(400).json({ error: 'include_hp must be a boolean (true or false).' });
    }

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      await client.query('UPDATE sewtech_batches SET include_hp=? WHERE id=?', [include_hp, id]);

      // Recompute operators and Mc Hrs for all rows in batch
      const { rows } = await client.query('SELECT * FROM sewtech_rows WHERE batch_id=?', [id]);
      for (const r of rows) {
        engine.calculateRowMetrics(r, include_hp);
        await client.query(
          `UPDATE sewtech_rows
           SET operators_final=?, mc_hrs=?, earn_minutes=?, effi=?
           WHERE id=?`,
          [r.operators_final, r.mc_hrs, r.earn_minutes, r.effi, r.id]
        );
      }

      await client.query(
        `INSERT INTO sewtech_change_log (batch_id, category, details)
         VALUES (?, 'hp_decision', ?)`,
        [id, { include_hp, description: include_hp ? 'Included OP + HP' : 'Operators only (OP)' }]
      );

      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    res.json({ ok: true, include_hp });
  } catch (err) {
    console.error('Sewtech hp-decision error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/sewtech/:id/working-days  { working_days }
router.post('/:id/working-days', async (req, res) => {
  try {
    const { id } = req.params;
    const { working_days } = req.body;

    const days = Number(working_days);
    if (!days || isNaN(days) || days < 1) {
      return res.status(400).json({ error: 'working_days must be a valid positive number.' });
    }

    await pool.query(
      `UPDATE sewtech_batches
       SET working_days=?, working_days_source='user_confirmed'
       WHERE id=?`,
      [days, id]
    );

    // If batch is already finalized, update summary
    const batch = (await pool.query('SELECT * FROM sewtech_batches WHERE id=?', [id])).rows[0];
    if (batch.status === 'finalized') {
      await pool.query(
        `UPDATE sewtech_final_summary
         SET avg_prod_per_day = sum_production / ?,
             avg_optrs_per_day = sum_operators / ?
         WHERE batch_id=?`,
        [days, days, id]
      );
    }

    await pool.query(
      `INSERT INTO sewtech_change_log (batch_id, category, details)
       VALUES (?, 'working_days_edited', ?)`,
      [id, { confirmed_working_days: days }]
    );

    res.json({ ok: true, working_days: days });
  } catch (err) {
    console.error('Sewtech working-days error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/sewtech/:id/upload-modified (multipart "file")
router.post('/:id/upload-modified', upload.single('file'), async (req, res) => {
  try {
    const { id } = req.params;
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

    const parsedRows = excelIO.parseModifiedMergedSheet(req.file.buffer);
    if (parsedRows.length === 0) {
      return res.status(400).json({ error: 'No valid rows found in uploaded modified sheet.' });
    }

    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      await client.query('DELETE FROM sewtech_rows WHERE batch_id=?', [id]);

      for (const row of parsedRows) {
        await client.query(
          `INSERT INTO sewtech_rows (
             batch_id, work_date, line, unit, operators_op, helpers_hp, operators_final,
             reg_hours, ot_man, ot_hr, raw_buyer, buyer_group, is_pending, pending_reason,
             style, item_no, order_qty, smv, earn_minutes, production, effi, mc_hrs
           ) VALUES (
             ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
           )`,
          [
            id, row.work_date || new Date().toISOString().slice(0, 10), row.line, row.unit,
            row.operators_final, 0, row.operators_final, row.reg_hours, row.ot_man, row.ot_hr,
            row.raw_buyer, row.buyer_group, false, null, row.style, row.item_no, row.order_qty,
            row.smv, row.earn_minutes, row.production, row.effi, row.mc_hrs,
          ]
        );
      }

      await client.query(
        `UPDATE sewtech_batches SET status='modified_uploaded' WHERE id=?`,
        [id]
      );

      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    res.json({ ok: true, row_count: parsedRows.length });
  } catch (err) {
    console.error('Sewtech upload-modified error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/sewtech/:id/finalize
router.post('/:id/finalize', async (req, res) => {
  try {
    const { id } = req.params;
    const batchRes = await pool.query('SELECT * FROM sewtech_batches WHERE id=?', [id]);
    const batch = batchRes.rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    // Check for pending items
    const pendingRes = await pool.query(
      `SELECT COUNT(*) AS count FROM sewtech_rows WHERE batch_id=? AND is_pending=TRUE`,
      [id]
    );
    if (Number(pendingRes.rows[0].count) > 0) {
      return res.status(400).json({
        error: 'Cannot finalize: there are unresolved pending buyers in this batch. Please map them first.',
      });
    }

    const hpCountRes = await pool.query(
      `SELECT COUNT(*) AS count FROM sewtech_rows WHERE batch_id=? AND helpers_hp > 0`,
      [id]
    );
    if (Number(hpCountRes.rows[0].count) > 0 && batch.include_hp == null) {
      return res.status(400).json({
        error: 'Cannot finalize: Helper (HP) decision is pending. Please choose OP only or OP+HP.',
      });
    }

    const { rows } = await pool.query('SELECT * FROM sewtech_rows WHERE batch_id=?', [id]);
    const { summary, grandTotal } = engine.buildFinalSummary(rows, batch.working_days);

    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      await client.query('DELETE FROM sewtech_final_summary WHERE batch_id=?', [id]);

      for (const item of summary) {
        await client.query(
          `INSERT INTO sewtech_final_summary (
             batch_id, buyer_group, sum_production, sum_operators, avg_prod_per_day, avg_optrs_per_day
           ) VALUES (?, ?, ?, ?, ?, ?)`,
          [id, item.buyer_group, item.sum_production, item.sum_operators, item.avg_prod_per_day, item.avg_optrs_per_day]
        );
      }

      await client.query(
        `UPDATE sewtech_batches SET status='finalized', finalized_at=CURRENT_TIMESTAMP WHERE id=?`,
        [id]
      );

      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    res.json({ ok: true, summary, grand_total: grandTotal });
  } catch (err) {
    console.error('Sewtech finalize error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/sewtech/:id/final-export
router.get('/:id/final-export', async (req, res) => {
  try {
    const { id } = req.params;
    const batchRes = await pool.query('SELECT * FROM sewtech_batches WHERE id=?', [id]);
    const batch = batchRes.rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    let summaryRows = (
      await pool.query('SELECT * FROM sewtech_final_summary WHERE batch_id=? ORDER BY buyer_group', [id])
    ).rows;

    // If not yet written to final_summary table, compute dynamically
    if (summaryRows.length === 0) {
      const { rows } = await pool.query('SELECT * FROM sewtech_rows WHERE batch_id=?', [id]);
      const resSum = engine.buildFinalSummary(rows, batch.working_days);
      summaryRows = resSum.summary;
    }

    const buffer = await excelIO.writeFinalSummarySheet(summaryRows, {
      startDate: batch.start_date,
      endDate: batch.end_date,
      weekLabel: batch.week_label,
      workingDays: batch.working_days,
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Sewtech_Final_Summary_Batch_${id}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error('Sewtech final-export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/sewtech (batch history list)
router.get('/', async (req, res) => {
  try {
    const { rows } = await pool.query('SELECT * FROM sewtech_batches ORDER BY id DESC LIMIT 50');
    res.json(rows);
  } catch (err) {
    console.error('Sewtech batch list error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Master data endpoints: Buyer groups
router.get('/master/buyer-groups', async (req, res) => {
  try {
    const { rows } = await pool.query('SELECT * FROM sewtech_buyer_groups ORDER BY name');
    res.json(rows);
  } catch (err) {
    console.error('Sewtech buyer groups error:', err);
    res.status(500).json({ error: err.message });
  }
});

router.post('/master/buyer-groups', async (req, res) => {
  try {
    const name = engine.normalizeText(req.body.name);
    if (!name) return res.status(400).json({ error: 'name is required' });
    await pool.query(
      `INSERT INTO sewtech_buyer_groups (name) VALUES (?)
       ON DUPLICATE KEY UPDATE name=VALUES(name)`,
      [name]
    );
    const { rows } = await pool.query('SELECT * FROM sewtech_buyer_groups WHERE name=?', [name]);
    res.json(rows[0]);
  } catch (err) {
    console.error('Sewtech add buyer group error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Master data endpoints: Buyer mappings
router.get('/master/buyer-map', async (req, res) => {
  try {
    const { rows } = await pool.query(`
      SELECT m.id, m.raw_buyer_norm, m.raw_buyer_display, m.buyer_group_id, g.name AS buyer_group, m.created_at
      FROM sewtech_buyer_map m
      JOIN sewtech_buyer_groups g ON m.buyer_group_id = g.id
      ORDER BY m.raw_buyer_display
    `);
    res.json(rows);
  } catch (err) {
    console.error('Sewtech buyer map error:', err);
    res.status(500).json({ error: err.message });
  }
});

router.post('/master/buyer-map', async (req, res) => {
  try {
    const { raw_buyer, buyer_group_id, buyer_group } = req.body;
    const rawDisplay = engine.normalizeText(raw_buyer);
    const rawNorm = engine.normalizeKey(raw_buyer);

    if (!rawDisplay) return res.status(400).json({ error: 'raw_buyer is required' });

    let groupId = buyer_group_id;
    if (!groupId && buyer_group) {
      const gName = engine.normalizeText(buyer_group);
      await pool.query(
        `INSERT INTO sewtech_buyer_groups (name) VALUES (?)
         ON DUPLICATE KEY UPDATE name=VALUES(name)`,
        [gName]
      );
      const gRes = await pool.query('SELECT id FROM sewtech_buyer_groups WHERE name=?', [gName]);
      groupId = gRes.rows[0].id;
    }

    if (!groupId) return res.status(400).json({ error: 'buyer_group_id or buyer_group is required' });

    await pool.query(
      `INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id, updated_at)
       VALUES (?, ?, ?, CURRENT_TIMESTAMP)
       ON DUPLICATE KEY UPDATE
         buyer_group_id=VALUES(buyer_group_id),
         raw_buyer_display=VALUES(raw_buyer_display),
         updated_at=CURRENT_TIMESTAMP`,
      [rawNorm, rawDisplay, groupId]
    );

    const { rows } = await pool.query('SELECT * FROM sewtech_buyer_map WHERE raw_buyer_norm=?', [rawNorm]);
    res.json(rows[0]);
  } catch (err) {
    console.error('Sewtech add buyer map error:', err);
    res.status(500).json({ error: err.message });
  }
});

router.delete('/master/buyer-map/:id', async (req, res) => {
  try {
    await pool.query('DELETE FROM sewtech_buyer_map WHERE id=?', [req.params.id]);
    res.json({ ok: true });
  } catch (err) {
    console.error('Sewtech delete buyer map error:', err);
    res.status(500).json({ error: err.message });
  }
});

// DELETE /api/sewtech/:id — delete batch and cascaded rows
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query('DELETE FROM sewtech_batches WHERE id=?', [id]);
    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Batch not found' });
    }
    res.json({ ok: true, deleted_id: Number(id) });
  } catch (err) {
    console.error('Sewtech delete batch error:', err);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
