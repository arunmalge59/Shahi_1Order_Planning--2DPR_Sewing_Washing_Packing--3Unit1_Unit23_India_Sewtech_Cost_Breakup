-- Shahi Textile Automation — MySQL schema.

CREATE TABLE IF NOT EXISTS brand_buyer_map (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  brand VARCHAR(255) NOT NULL,
  buyer_group VARCHAR(255) NOT NULL,
  item_pattern VARCHAR(255) NULL,
  item_pattern_key VARCHAR(255) GENERATED ALWAYS AS (COALESCE(item_pattern, '')) STORED,
  is_denim_split BOOLEAN NOT NULL DEFAULT FALSE,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_brand_buyer_pattern (brand, item_pattern),
  UNIQUE KEY uq_brand_buyer_pattern_normalized (brand, item_pattern_key)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

-- Per-item overrides, for brands that split into more than one buyer group.
CREATE TABLE IF NOT EXISTS item_buyer_override (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  brand VARCHAR(255) NOT NULL,
  item VARCHAR(255) NOT NULL,
  buyer_group TEXT NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_item_buyer_brand_item (brand, item)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

-- Standard SAM-per-piece master, used to fix rows where raw SAM = 0.
CREATE TABLE IF NOT EXISTS sam_master (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  item VARCHAR(255) NOT NULL,
  shahi_style TEXT NULL,
  standard_sam_per_piece DECIMAL(65, 30) NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uq_sam_master_item (item)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

-- One row per weekly upload.
CREATE TABLE IF NOT EXISTS batches (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  week_label TEXT NOT NULL,
  working_days INT NOT NULL DEFAULT 5,
  raw_filename TEXT NOT NULL,
  status VARCHAR(64) NOT NULL DEFAULT 'raw_uploaded',
  uploaded_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  finalized_at DATETIME NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

-- Every row of the working (modification) sheet, after cleaning.
CREATE TABLE IF NOT EXISTS batch_rows (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id INT NOT NULL,
  raw_row_no INT NULL,
  facility TEXT NULL,
  pch TEXT NULL,
  act_ctrl_obj TEXT NULL,
  sales_person TEXT NULL,
  item_responsible TEXT NULL,
  item TEXT NULL,
  item_group TEXT NULL,
  type TEXT NULL,
  item_category TEXT NULL,
  shahi_style TEXT NULL,
  unit TEXT NULL,
  brand VARCHAR(255) NOT NULL,
  raw_buyer TEXT NULL,
  resolved_buyer TEXT NULL,
  is_new_brand BOOLEAN NOT NULL DEFAULT FALSE,
  order_qty DECIMAL(65, 30) NULL,
  production_qty DECIMAL(65, 30) NULL,
  mcs DECIMAL(65, 30) NULL,
  budgeted_cost DECIMAL(65, 30) NULL,
  actual_cost DECIMAL(65, 30) NULL,
  ot_cost DECIMAL(65, 30) NULL,
  total_cost DECIMAL(65, 30) NULL,
  cost_per_pcs DECIMAL(65, 30) NULL,
  sam DECIMAL(65, 30) NULL,
  budgeted_cost_per_sam DECIMAL(65, 30) NULL,
  actual_cost_per_sam DECIMAL(65, 30) NULL,
  budgeted_efficiency DECIMAL(65, 30) NULL,
  actual_efficiency DECIMAL(65, 30) NULL,
  fob_gmt_inr DECIMAL(65, 30) NULL,
  currency TEXT NULL,
  fob_gmt_fc DECIMAL(65, 30) NULL,
  currency_conversion DECIMAL(65, 30) NULL,
  was_sam_zero_fixed BOOLEAN NOT NULL DEFAULT FALSE,
  was_deleted BOOLEAN NOT NULL DEFAULT FALSE,
  delete_reason TEXT NULL,
  CONSTRAINT fk_batch_rows_batch FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE,
  KEY idx_batch_rows_batch (batch_id),
  KEY idx_batch_rows_brand (brand)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

-- Full audit log of every automated decision made on a batch.
CREATE TABLE IF NOT EXISTS change_log (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id INT NOT NULL,
  category TEXT NOT NULL,
  item TEXT NULL,
  brand TEXT NULL,
  raw_row_no INT NULL,
  details JSON NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_change_log_batch FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

-- Final pivot-by-buyer summary, one set of rows per finalized batch.
CREATE TABLE IF NOT EXISTS final_summary (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id INT NOT NULL,
  buyer_group VARCHAR(255) NOT NULL,
  sum_production_qty DECIMAL(65, 30) NOT NULL,
  sum_mcs DECIMAL(65, 30) NOT NULL,
  avg_prod_per_day DECIMAL(65, 30) NOT NULL,
  avg_optrs_per_day DECIMAL(65, 30) NOT NULL,
  CONSTRAINT fk_final_summary_batch FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE,
  UNIQUE KEY uq_final_summary_batch_buyer (batch_id, buyer_group)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;
