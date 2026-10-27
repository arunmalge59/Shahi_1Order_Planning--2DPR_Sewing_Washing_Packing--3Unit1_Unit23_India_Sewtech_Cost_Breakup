const express = require('express');
const multer = require('multer');
const pool = require('../db/pool');
const engine = require('../services/engine');
const excelIO = require('../services/excelIO');

const router = express.Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

async function loadBrandMap() {
  const { rows } = await pool.query('SELECT brand, buyer_group, is_denim_split FROM brand_buyer_map WHERE item_pattern IS NULL');
  const map = new Map();
  for (const r of rows) map.set(r.brand, { buyer_group: r.buyer_group, is_denim_split: r.is_denim_split });
  return map;
}

async function loadOverrideMap() {
  const { rows } = await pool.query('SELECT brand, item, buyer_group FROM item_buyer_override');
  const map = new Map();
  for (const r of rows) map.set(`${r.brand}||${r.item}`, r.buyer_group);
  return map;
}

async function loadSamMasterMap() {
  const { rows } = await pool.query('SELECT item, standard_sam_per_piece FROM sam_master');
  const map = new Map();
  for (const r of rows) map.set(r.item, Number(r.standard_sam_per_piece));
  return map;
}

// POST /api/batches/upload-raw  (multipart file field: "file")
router.post('/upload-raw', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded (expected field "file").' });

    const weekLabel = req.body.week_label ? String(req.body.week_label).trim() : null;
    const workingDays = Number(req.body.working_days);

    if (!weekLabel) {
      return res.status(400).json({ error: 'Week label is mandatory.' });
    }
    if (!workingDays || isNaN(workingDays) || workingDays < 1) {
      return res.status(400).json({ error: 'Working days this week must be a valid number (e.g. 5, 6, 7).' });
    }

    const rawRows = excelIO.parseRawWorkbook(req.file.buffer, req.file.originalname);
    excelIO.coerceNumbers(rawRows);

    const [brandMap, overrideMap, samMasterMap] = await Promise.all([
      loadBrandMap(), loadOverrideMap(), loadSamMasterMap(),
    ]);

    const { rows, changeLog } = engine.runPipeline({ rawRows, brandMap, overrideMap, samMasterMap, workingDays });

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      const batchRes = await client.query(
        `INSERT INTO batches (week_label, working_days, raw_filename, status)
         VALUES (?, ?, ?, 'raw_uploaded')`,
        [weekLabel || `Uploaded ${new Date().toISOString().slice(0, 10)}`, workingDays, req.file.originalname]
      );
      const batchId = batchRes.insertId;

      for (const row of rows) {
        await client.query(
          `INSERT INTO batch_rows (
             batch_id, raw_row_no, facility, pch, act_ctrl_obj, sales_person, item_responsible,
             item, item_group, type, item_category, shahi_style, unit, brand, raw_buyer,
             resolved_buyer, is_new_brand, order_qty, production_qty, mcs,
             budgeted_cost, actual_cost, ot_cost, total_cost, cost_per_pcs, sam,
             budgeted_cost_per_sam, actual_cost_per_sam, budgeted_efficiency, actual_efficiency,
             fob_gmt_inr, currency, fob_gmt_fc, currency_conversion, was_sam_zero_fixed,
             was_deleted, delete_reason
           ) VALUES (
             ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,
             ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
           )`,
          [
            batchId, row.raw_row_no || null, row.facility, row.pch, row.act_ctrl_obj, row.sales_person, row.item_responsible,
            row.item, row.item_group, row.type, row.item_category, row.shahi_style, row.unit, row.brand, row.raw_buyer,
            row.resolved_buyer, !!row.is_new_brand, row.order_qty, row.production_qty, row.mcs,
            row.budgeted_cost, row.actual_cost, row.ot_cost, row.total_cost, row.cost_per_pcs, row.sam,
            row.budgeted_cost_per_sam, row.actual_cost_per_sam, row.budgeted_efficiency, row.actual_efficiency,
            row.fob_gmt_inr, row.currency, row.fob_gmt_fc, row.currency_conversion, !!row.was_sam_zero_fixed,
            !!row.was_deleted, row.delete_reason || null,
          ]
        );
      }

      for (const entry of changeLog) {
        await client.query(
          `INSERT INTO change_log (batch_id, category, item, brand, raw_row_no, details) VALUES (?,?,?,?,?,?)`,
          [batchId, entry.category, entry.item || null, entry.brand || null, entry.raw_row_no || entry.details?.raw_row_no || null, entry.details || {}]
        );
      }

      await client.query('COMMIT');
      res.json({ batch_id: batchId, row_count: rows.length, change_log_count: changeLog.length });
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/batches/:id/report — "what was done" summary, shown right after upload
router.get('/:id/report', async (req, res) => {
  try {
    const { id } = req.params;
    const batch = (await pool.query('SELECT * FROM batches WHERE id=?', [id])).rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    const logRows = (await pool.query(
      `SELECT category, item, brand, raw_row_no, details, created_at FROM change_log WHERE batch_id=? ORDER BY id`, [id]
    )).rows;

    const grouped = {
      row_deleted: [],
      new_brand: [],
      sam_eff_fixed: [],
      buyer_split_needed: [],
      forward_fill_count: 0,
    };
    for (const l of logRows) {
      if (l.category === 'forward_fill') grouped.forward_fill_count += 1;
      else if (grouped[l.category]) grouped[l.category].push(l);
    }

    const totals = (await pool.query(
      `SELECT
         COUNT(IF(NOT was_deleted, 1, NULL)) AS kept_rows,
         COUNT(IF(was_deleted, 1, NULL)) AS deleted_rows,
         COUNT(IF(is_new_brand, 1, NULL)) AS new_brand_rows,
         COUNT(IF(was_sam_zero_fixed, 1, NULL)) AS sam_fixed_rows
       FROM batch_rows WHERE batch_id=?`, [id]
    )).rows[0];

    let finalPivot = null;
    if (batch.status === 'finalized') {
      finalPivot = (await pool.query('SELECT * FROM final_summary WHERE batch_id=?', [id])).rows;
    }

    res.json({ batch, totals, finalPivot, ...grouped });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/batches/:id/export-modification — cleaned xlsx, new brands highlighted red
router.get('/:id/export-modification', async (req, res) => {
  try {
    const { id } = req.params;
    const batch = (await pool.query('SELECT * FROM batches WHERE id=?', [id])).rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    const rows = (await pool.query('SELECT * FROM batch_rows WHERE batch_id=? ORDER BY id', [id])).rows;
    const buffer = await excelIO.writeModificationSheet(rows, { weekLabel: batch.week_label });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="modification_batch_${id}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/batches/:id/assign-buyer  { brand, item (optional), buyer_group, save_to_master }
// Used when the user resolves a "new_brand" or "buyer_split_needed" flag from the UI.
router.post('/:id/assign-buyer', async (req, res) => {
  try {
    const { id } = req.params;
    const { brand, item, buyer_group, save_to_master } = req.body;
    if (!brand || !buyer_group) return res.status(400).json({ error: 'brand and buyer_group are required' });

    if (item) {
      await pool.query(
        `UPDATE batch_rows SET resolved_buyer=?, is_new_brand=FALSE WHERE batch_id=? AND brand=? AND item=?`,
        [buyer_group, id, brand, item]
      );
      if (save_to_master) {
        await pool.query(
          `INSERT INTO item_buyer_override (brand, item, buyer_group) VALUES (?,?,?)
           ON DUPLICATE KEY UPDATE buyer_group=VALUES(buyer_group)`,
          [brand, item, buyer_group]
        );
      }
    } else {
      await pool.query(
        `UPDATE batch_rows SET resolved_buyer=?, is_new_brand=FALSE WHERE batch_id=? AND brand=?`,
        [buyer_group, id, brand]
      );
      if (save_to_master) {
        await pool.query(
          `INSERT INTO brand_buyer_map (brand, buyer_group, item_pattern, is_denim_split) VALUES (?,?,NULL,FALSE)
           ON DUPLICATE KEY UPDATE buyer_group=VALUES(buyer_group), updated_at=CURRENT_TIMESTAMP`,
          [brand, buyer_group]
        );
      }
    }
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/batches/:id/finalize — recompute pivot from (possibly user-edited) batch_rows and lock it in
router.post('/:id/finalize', async (req, res) => {
  try {
    const { id } = req.params;
    const batch = (await pool.query('SELECT * FROM batches WHERE id=?', [id])).rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    const rows = (await pool.query('SELECT * FROM batch_rows WHERE batch_id=?', [id])).rows;
    const pivot = engine.pivotByBuyer(rows, batch.working_days);

    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      await client.query('DELETE FROM final_summary WHERE batch_id=?', [id]);
      for (const g of pivot) {
        await client.query(
          `INSERT INTO final_summary (batch_id, buyer_group, sum_production_qty, sum_mcs, avg_prod_per_day, avg_optrs_per_day)
           VALUES (?,?,?,?,?,?)`,
          [id, g.buyer_group, g.sum_production_qty, g.sum_mcs, g.avg_prod_per_day, g.avg_optrs_per_day]
        );
      }
      await client.query(`UPDATE batches SET status='finalized', finalized_at=CURRENT_TIMESTAMP WHERE id=?`, [id]);
      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }

    res.json({ ok: true, pivot });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/batches/:id/final-export — final buyer-wise summary as xlsx
router.get('/:id/final-export', async (req, res) => {
  try {
    const { id } = req.params;
    const batch = (await pool.query('SELECT * FROM batches WHERE id=?', [id])).rows[0];
    if (!batch) return res.status(404).json({ error: 'Batch not found' });

    const pivot = (await pool.query('SELECT * FROM final_summary WHERE batch_id=?', [id])).rows;
    const buffer = await excelIO.writeFinalSummarySheet(pivot, { weekLabel: batch.week_label, workingDays: batch.working_days });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="final_summary_batch_${id}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/batches — list, newest first
router.get('/', async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM batches ORDER BY id DESC LIMIT 50');
  res.json(rows);
});

// GET /api/batches/:id/rows — full row-level data (for the review grid)
router.get('/:id/rows', async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM batch_rows WHERE batch_id=? ORDER BY id', [req.params.id]);
  res.json(rows);
});

// DELETE /api/batches/:id — delete batch and cascaded rows
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query('DELETE FROM batches WHERE id=?', [id]);
    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Batch not found' });
    }
    res.json({ ok: true, deleted_id: Number(id) });
  } catch (err) {
    console.error('Delete batch error:', err);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
