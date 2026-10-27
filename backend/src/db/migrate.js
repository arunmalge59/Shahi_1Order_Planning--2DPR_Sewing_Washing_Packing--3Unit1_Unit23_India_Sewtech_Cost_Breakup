require('dotenv').config();

const path = require('path');
const mysql = require('mysql2/promise');

const dbName = process.env.DB_NAME || 'shahi_unified_operations_db';
if (!/^[A-Za-z0-9_]+$/.test(dbName)) {
  throw new Error('DB_NAME may contain only letters, numbers, and underscores.');
}

async function migrate() {
  const bootstrap = await mysql.createConnection({
    host: process.env.DB_HOST || 'localhost',
    port: Number(process.env.DB_PORT || 3306),
    user: process.env.DB_USER || 'root',
    password: process.env.DB_PASSWORD || ''
  });

  try {
    await bootstrap.query(
      `CREATE DATABASE IF NOT EXISTS \`${dbName}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci`
    );
  } finally {
    await bootstrap.end();
  }
const executeSqlFile = require('../../modules/cost/src/db/executeSqlFile');
const costPool = require('../../modules/cost/src/db/pool');
const productionDb = require('../../modules/production/src/db');
const planningDb = require('../../modules/planning/src/db');

  try {
    await executeSqlFile(productionDb.pool, path.resolve(__dirname, '../../modules/production/schema.sql'));
await executeSqlFile(costPool, path.resolve(__dirname, '../../modules/cost/src/db/schema.sql'));
await executeSqlFile(costPool, path.resolve(__dirname, '../../modules/cost/src/db/sewtech_schema.sql'));
    await planningDb.initDb();
    await productionDb.initializeDatabase();
    console.log(`Database schemas ensured in ${dbName}. Existing rows are not deleted.`);
  } finally {
    await Promise.all([
      costPool.end(),
      productionDb.pool.end(),
      planningDb.pool.end()
    ]);
  }
}

migrate().catch((error) => {
  console.error('Database migration failed:', error);
  process.exitCode = 1;
});
