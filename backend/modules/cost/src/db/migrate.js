const path = require('path');
const pool = require('./pool');
const executeSqlFile = require('./executeSqlFile');

async function migrate() {
  try {
    await executeSqlFile(pool, path.join(__dirname, 'schema.sql'));
    await executeSqlFile(pool, path.join(__dirname, 'sewtech_schema.sql'));
    console.log('Migration complete: India and Sewtech MySQL tables ensured.');
  } finally {
    await pool.end();
  }
}

migrate().catch((err) => {
  console.error('Migration failed:', err);
  process.exitCode = 1;
});
