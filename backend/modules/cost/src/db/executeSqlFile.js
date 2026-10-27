const fs = require('fs');

async function executeSqlFile(pool, filePath) {
  const sql = fs.readFileSync(filePath, 'utf8').replace(/^\s*--.*$/gm, '');
  const statements = sql.split(';').map((statement) => statement.trim()).filter(Boolean);
  for (const statement of statements) {
    await pool.query(statement);
  }
}

module.exports = executeSqlFile;
