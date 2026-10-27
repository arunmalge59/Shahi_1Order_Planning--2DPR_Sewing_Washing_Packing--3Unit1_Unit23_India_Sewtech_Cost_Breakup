const path = require('path');
const pool = require('./pool');
const executeSqlFile = require('./executeSqlFile');

async function seed() {
  try {
    await executeSqlFile(pool, path.join(__dirname, '..', '..', 'seed', 'seed_brand_buyer.sql'));
    await executeSqlFile(pool, path.join(__dirname, '..', '..', 'seed', 'seed_sewtech.sql'));
    console.log('Seed complete: India and Sewtech seed data populated.');
  } finally {
    await pool.end();
  }
}

seed().catch((err) => {
  console.error('Seed failed:', err);
  process.exitCode = 1;
});
