const mysql = require('mysql2/promise');
require('dotenv').config();

const mysqlPool = mysql.createPool({
  host: process.env.DB_HOST || process.env.MYSQL_HOST || 'localhost',
  port: Number(process.env.DB_PORT || process.env.MYSQL_PORT || 3306),
  user: process.env.DB_USER || process.env.MYSQL_USER || 'root',
  password: process.env.DB_PASSWORD ?? process.env.MYSQL_PASSWORD ?? '',
  database: process.env.DB_NAME || process.env.MYSQL_DATABASE || 'shahi_unified_operations_db',
  charset: 'utf8mb4',
  waitForConnections: true,
  connectionLimit: Number(process.env.MYSQL_CONNECTION_LIMIT || 10),
  queueLimit: 0,
  timezone: 'Z',
  dateStrings: ['DATE'],
  typeCast(field, next) {
    if (field.type === 'TINY' && field.length === 1) {
      const value = field.string();
      return value === null ? null : value === '1';
    }
    return next();
  },
});

function normalizeParams(params) {
  return params.map((value) => {
    if (value === undefined) return null;
    if (value && typeof value === 'object' && !Buffer.isBuffer(value) && !(value instanceof Date)) {
      return JSON.stringify(value);
    }
    return value;
  });
}

async function queryWith(connection, sql, params = []) {
  const [result] = await connection.execute(sql, normalizeParams(params));
  if (Array.isArray(result)) {
    return { rows: result, rowCount: result.length };
  }
  return {
    rows: [],
    rowCount: result.affectedRows,
    insertId: result.insertId,
  };
}

const pool = {
  query: (sql, params) => queryWith(mysqlPool, sql, params),
  async connect() {
    const connection = await mysqlPool.getConnection();
    return {
      async query(sql, params) {
        const command = sql.trim().toUpperCase();
        if (command === 'BEGIN' || command === 'START TRANSACTION') {
          await connection.beginTransaction();
          return { rows: [], rowCount: 0 };
        }
        if (command === 'COMMIT') {
          await connection.commit();
          return { rows: [], rowCount: 0 };
        }
        if (command === 'ROLLBACK') {
          await connection.rollback();
          return { rows: [], rowCount: 0 };
        }
        return queryWith(connection, sql, params);
      },
      release: () => connection.release(),
    };
  },
  end: () => mysqlPool.end(),
};

module.exports = pool;
