const mysql = require('mysql2/promise');

const pool = mysql.createPool({
  host: process.env.DB_HOST || 'localhost',
  port: Number(process.env.DB_PORT || 3306),
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'shahi_unified_operations_db',
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
  typeCast(field, next) {
    if (field.type === 'TINY' && field.length === 1) {
      return field.string() === '1';
    }
    return next();
  },
});

async function initDb() {
  // Creates the tables the app needs if they don't already exist.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS comparison_runs (
      id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
      older_file_name TEXT NOT NULL,
      newer_file_name TEXT NOT NULL,
      total_styles INTEGER NOT NULL,
      changed_styles INTEGER NOT NULL,
      qty_changed INTEGER NOT NULL,
      pcd_changed INTEGER NOT NULL,
      ex_factory_changed INTEGER NOT NULL,
      created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB;
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS comparison_style_changes (
      id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
      run_id INT,
      style TEXT NOT NULL,
      july_qty DECIMAL(65,30),
      august_qty DECIMAL(65,30),
      qty_delta DECIMAL(65,30),
      pcd_changed BOOLEAN,
      ex_factory_changed BOOLEAN,
      status TEXT,
      CONSTRAINT comparison_style_changes_run_id_fkey
        FOREIGN KEY (run_id) REFERENCES comparison_runs(id) ON DELETE CASCADE
    ) ENGINE=InnoDB;
  `);
}

module.exports = { pool, initDb };
