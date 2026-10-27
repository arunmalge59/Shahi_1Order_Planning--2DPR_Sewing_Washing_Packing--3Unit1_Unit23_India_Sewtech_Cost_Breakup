const express = require('express');
const pool = require('../db/pool');

const router = express.Router();

// ---- Brand -> Buyer group mapping ----
router.get('/brand-buyer', async (req, res) => {
  const { rows } = await pool.query(
    'SELECT id, brand, buyer_group, item_pattern, is_denim_split, created_at, updated_at FROM brand_buyer_map ORDER BY brand'
  );
  res.json(rows);
});

router.post('/brand-buyer', async (req, res) => {
  const { brand, buyer_group, is_denim_split } = req.body;
  if (!brand || !buyer_group) return res.status(400).json({ error: 'brand and buyer_group required' });
  await pool.query(
    `INSERT INTO brand_buyer_map (brand, buyer_group, item_pattern, is_denim_split)
     VALUES (?, ?, NULL, ?)
     ON DUPLICATE KEY UPDATE buyer_group=VALUES(buyer_group), is_denim_split=VALUES(is_denim_split), updated_at=CURRENT_TIMESTAMP`,
    [brand, buyer_group, !!is_denim_split]
  );
  const { rows } = await pool.query(
    `SELECT id, brand, buyer_group, item_pattern, is_denim_split, created_at, updated_at
     FROM brand_buyer_map WHERE brand=? AND item_pattern IS NULL ORDER BY id DESC LIMIT 1`,
    [brand]
  );
  res.json(rows[0]);
});

router.delete('/brand-buyer/:id', async (req, res) => {
  await pool.query('DELETE FROM brand_buyer_map WHERE id=?', [req.params.id]);
  res.json({ ok: true });
});

// ---- Item-level buyer overrides (for split brands e.g. ADIDAS, H&M) ----
router.get('/item-overrides', async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM item_buyer_override ORDER BY brand, item');
  res.json(rows);
});

router.post('/item-overrides', async (req, res) => {
  const { brand, item, buyer_group } = req.body;
  if (!brand || !item || !buyer_group) return res.status(400).json({ error: 'brand, item, buyer_group required' });
  await pool.query(
    `INSERT INTO item_buyer_override (brand, item, buyer_group) VALUES (?, ?, ?)
     ON DUPLICATE KEY UPDATE buyer_group=VALUES(buyer_group)`,
    [brand, item, buyer_group]
  );
  const { rows } = await pool.query(
    'SELECT * FROM item_buyer_override WHERE brand=? AND item=?',
    [brand, item]
  );
  res.json(rows[0]);
});

router.delete('/item-overrides/:id', async (req, res) => {
  await pool.query('DELETE FROM item_buyer_override WHERE id=?', [req.params.id]);
  res.json({ ok: true });
});

// ---- Standard SAM-per-piece master ----
router.get('/sam-master', async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM sam_master ORDER BY item');
  res.json(rows);
});

router.post('/sam-master', async (req, res) => {
  const { item, shahi_style, standard_sam_per_piece } = req.body;
  if (!item || standard_sam_per_piece == null) return res.status(400).json({ error: 'item and standard_sam_per_piece required' });
  await pool.query(
    `INSERT INTO sam_master (item, shahi_style, standard_sam_per_piece) VALUES (?, ?, ?)
     ON DUPLICATE KEY UPDATE standard_sam_per_piece=VALUES(standard_sam_per_piece), shahi_style=VALUES(shahi_style), updated_at=CURRENT_TIMESTAMP`,
    [item, shahi_style || null, standard_sam_per_piece]
  );
  const { rows } = await pool.query('SELECT * FROM sam_master WHERE item=?', [item]);
  res.json(rows[0]);
});

router.delete('/sam-master/:id', async (req, res) => {
  await pool.query('DELETE FROM sam_master WHERE id=?', [req.params.id]);
  res.json({ ok: true });
});

module.exports = router;
