-- Seed data for Sewtech automation

-- 1. Buyer groups
INSERT INTO sewtech_buyer_groups (name) VALUES
  ('Columbia'),
  ('Tommy'),
  ('Job Work')
ON DUPLICATE KEY UPDATE id = id;

-- 2. Buyer mappings (normalized raw buyer -> group)
INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id)
SELECT 'columbia', 'Columbia', id FROM sewtech_buyer_groups WHERE name = 'Columbia'
ON DUPLICATE KEY UPDATE buyer_group_id = VALUES(buyer_group_id);

INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id)
SELECT 'tommy', 'Tommy', id FROM sewtech_buyer_groups WHERE name = 'Tommy'
ON DUPLICATE KEY UPDATE buyer_group_id = VALUES(buyer_group_id);

INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id)
SELECT 'bd design', 'BD Design', id FROM sewtech_buyer_groups WHERE name = 'Job Work'
ON DUPLICATE KEY UPDATE buyer_group_id = VALUES(buyer_group_id);

INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id)
SELECT 'texas vintage', 'Texas Vintage', id FROM sewtech_buyer_groups WHERE name = 'Job Work'
ON DUPLICATE KEY UPDATE buyer_group_id = VALUES(buyer_group_id);

INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id)
SELECT 'total apperals', 'Total Apperals', id FROM sewtech_buyer_groups WHERE name = 'Job Work'
ON DUPLICATE KEY UPDATE buyer_group_id = VALUES(buyer_group_id);

INSERT INTO sewtech_buyer_map (raw_buyer_norm, raw_buyer_display, buyer_group_id)
SELECT 'versatile', 'Versatile', id FROM sewtech_buyer_groups WHERE name = 'Job Work'
ON DUPLICATE KEY UPDATE buyer_group_id = VALUES(buyer_group_id);

-- 3. Unit mapping
INSERT INTO sewtech_unit_map (title_number, unit_code) VALUES
  (10, 'B10'),
  (11, 'B11')
ON DUPLICATE KEY UPDATE id = id;
