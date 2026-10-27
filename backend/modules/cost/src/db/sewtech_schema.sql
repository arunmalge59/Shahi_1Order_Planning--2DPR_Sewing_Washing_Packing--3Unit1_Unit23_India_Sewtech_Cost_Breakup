-- Sewtech Automation — additive MySQL tables.

CREATE TABLE IF NOT EXISTS sewtech_buyer_groups (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  name VARCHAR(255) NOT NULL UNIQUE,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

CREATE TABLE IF NOT EXISTS sewtech_buyer_map (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  raw_buyer_norm VARCHAR(255) NOT NULL UNIQUE,
  raw_buyer_display TEXT NOT NULL,
  buyer_group_id INT NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_sewtech_buyer_map_group FOREIGN KEY (buyer_group_id) REFERENCES sewtech_buyer_groups(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

CREATE TABLE IF NOT EXISTS sewtech_unit_map (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  title_number INT NOT NULL UNIQUE,
  unit_code TEXT NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

CREATE TABLE IF NOT EXISTS sewtech_batches (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  week_label TEXT NOT NULL,
  start_date DATE NOT NULL,
  end_date DATE NOT NULL,
  working_days INT NOT NULL DEFAULT 6,
  working_days_source VARCHAR(64) NOT NULL DEFAULT 'auto',
  raw_filename TEXT NOT NULL,
  raw_file LONGBLOB NULL,
  include_hp BOOLEAN NULL,
  status VARCHAR(64) NOT NULL DEFAULT 'raw_uploaded',
  uploaded_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  finalized_at DATETIME NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

CREATE TABLE IF NOT EXISTS sewtech_rows (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id INT NOT NULL,
  work_date DATE NOT NULL,
  line TEXT NOT NULL,
  unit TEXT NOT NULL,
  operators_op DECIMAL(65, 30) NOT NULL DEFAULT 0,
  helpers_hp DECIMAL(65, 30) NOT NULL DEFAULT 0,
  operators_final DECIMAL(65, 30) NOT NULL DEFAULT 0,
  reg_hours DECIMAL(65, 30) NOT NULL DEFAULT 8,
  ot_man DECIMAL(65, 30) NOT NULL DEFAULT 0,
  ot_hr DECIMAL(65, 30) NOT NULL DEFAULT 0,
  raw_buyer TEXT NULL,
  buyer_group TEXT NULL,
  is_pending BOOLEAN NOT NULL DEFAULT FALSE,
  pending_reason TEXT NULL,
  style TEXT NULL,
  item_no TEXT NULL,
  order_qty DECIMAL(65, 30) NULL,
  smv DECIMAL(65, 30) NULL,
  earn_minutes DECIMAL(65, 30) NULL,
  production DECIMAL(65, 30) NOT NULL DEFAULT 0,
  effi DECIMAL(65, 30) NULL,
  mc_hrs DECIMAL(65, 30) NULL,
  source_tab TEXT NULL,
  source_row_no INT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_sewtech_rows_batch FOREIGN KEY (batch_id) REFERENCES sewtech_batches(id) ON DELETE CASCADE,
  KEY idx_sewtech_rows_batch (batch_id),
  KEY idx_sewtech_rows_date (work_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

CREATE TABLE IF NOT EXISTS sewtech_change_log (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id INT NOT NULL,
  category TEXT NOT NULL,
  details JSON NOT NULL DEFAULT (JSON_OBJECT()),
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_sewtech_change_log_batch FOREIGN KEY (batch_id) REFERENCES sewtech_batches(id) ON DELETE CASCADE,
  KEY idx_sewtech_log_batch (batch_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;

CREATE TABLE IF NOT EXISTS sewtech_final_summary (
  id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
  batch_id INT NOT NULL,
  buyer_group VARCHAR(255) NOT NULL,
  sum_production DECIMAL(65, 30) NOT NULL DEFAULT 0,
  sum_operators DECIMAL(65, 30) NOT NULL DEFAULT 0,
  avg_prod_per_day DECIMAL(65, 30) NOT NULL DEFAULT 0,
  avg_optrs_per_day DECIMAL(65, 30) NOT NULL DEFAULT 0,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT fk_sewtech_final_summary_batch FOREIGN KEY (batch_id) REFERENCES sewtech_batches(id) ON DELETE CASCADE,
  UNIQUE KEY uq_sewtech_summary_batch_buyer (batch_id, buyer_group)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_bin;
