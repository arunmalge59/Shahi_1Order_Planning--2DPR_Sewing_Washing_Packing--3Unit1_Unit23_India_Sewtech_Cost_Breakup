CREATE TABLE IF NOT EXISTS comparison_runs (
  id INT NOT NULL AUTO_INCREMENT,
  older_file_name TEXT NOT NULL,
  newer_file_name TEXT NOT NULL,
  total_styles INTEGER NOT NULL,
  changed_styles INTEGER NOT NULL,
  qty_changed INTEGER NOT NULL,
  pcd_changed INTEGER NOT NULL,
  ex_factory_changed INTEGER NOT NULL,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS comparison_style_changes (
  id INT NOT NULL AUTO_INCREMENT,
  run_id INT,
  style TEXT NOT NULL,
  july_qty DECIMAL(65,30),
  august_qty DECIMAL(65,30),
  qty_delta DECIMAL(65,30),
  pcd_changed BOOLEAN,
  ex_factory_changed BOOLEAN,
  status TEXT,
  PRIMARY KEY (id),
  CONSTRAINT comparison_style_changes_run_id_fkey
    FOREIGN KEY (run_id) REFERENCES comparison_runs(id) ON DELETE CASCADE
) ENGINE=InnoDB;
