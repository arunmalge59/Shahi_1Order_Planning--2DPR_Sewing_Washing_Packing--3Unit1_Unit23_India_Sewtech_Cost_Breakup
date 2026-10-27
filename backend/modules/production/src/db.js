const mysql = require('mysql2/promise');

const sslEnabled = process.env.DB_SSL !== undefined
  ? process.env.DB_SSL.toLowerCase() === 'true'
  : process.env.NODE_ENV === 'production' && process.env.DB_HOST !== 'localhost';

const pool = mysql.createPool({
  host: process.env.DB_HOST || 'localhost',
  port: Number(process.env.DB_PORT || 3306),
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME || 'shahi_unified_operations_db',
  ssl: sslEnabled ? { rejectUnauthorized: false } : undefined,
  waitForConnections: true,
  connectionLimit: Number(process.env.DB_POOL_MAX || 10),
  connectTimeout: Number(process.env.DB_CONNECTION_TIMEOUT_MS || 5000),
  idleTimeout: Number(process.env.DB_IDLE_TIMEOUT_MS || 30000),
  enableKeepAlive: true,
  timezone: 'Z',
});

pool.on('connection', (connection) => {
  connection.query("SET time_zone = '+00:00'", (err) => {
    if (err) {
      console.error('Could not set MySQL connection timezone to UTC:', formatError(err));
    }
  });
});

pool.on('error', (err) => {
  console.error('Unexpected MySQL error on idle connection:', formatError(err));
});

const RETRYABLE_CODES = new Set([
  'ER_CON_COUNT_ERROR', 'ER_TOO_MANY_USER_CONNECTIONS', 'PROTOCOL_CONNECTION_LOST',
  'PROTOCOL_ENQUEUE_AFTER_FATAL_ERROR', 'PROTOCOL_INCORRECT_PACKET_SEQUENCE',
  'ECONNRESET', 'ECONNREFUSED', 'ETIMEDOUT', 'EPIPE', 'ENOTFOUND',
]);
const MAX_RETRIES = 3;
let schemaReady = false;

function formatError(err) {
  return `${err.code ? `[${err.code}] ` : ''}${err.message || 'unknown error'}`;
}

function isRetryable(err) {
  return RETRYABLE_CODES.has(err.code);
}

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function queryWithRetry(text, values, label, retries = MAX_RETRIES) {
  for (let attempt = 0; ; attempt += 1) {
    try {
      const [result] = await pool.execute(text, values);
      return result;
    } catch (err) {
      if (!isRetryable(err) || attempt >= retries) {
        console.error(`MySQL ${label} failed after ${attempt + 1} attempt(s):`, formatError(err));
        throw err;
      }
      const wait = 250 * (2 ** attempt);
      console.warn(`MySQL ${label} connection failed; retrying in ${wait}ms:`, formatError(err));
      await delay(wait);
    }
  }
}

async function ensureSchema() {
  if (schemaReady) return;
  const columns = await queryWithRetry(
    `SELECT COLUMN_NAME
     FROM information_schema.COLUMNS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'jobs' AND COLUMN_NAME = 'request_id'`,
    [],
    'job schema column check'
  );
  if (columns.length === 0) {
    try {
      await queryWithRetry(
        'ALTER TABLE jobs ADD COLUMN request_id VARCHAR(255) NULL',
        [],
        'job schema migration'
      );
    } catch (err) {
      if (err.code !== 'ER_DUP_FIELDNAME') throw err;
    }
  }

  const indexes = await queryWithRetry(
    `SELECT INDEX_NAME
     FROM information_schema.STATISTICS
     WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'jobs'
       AND INDEX_NAME = 'idx_jobs_request_id' AND NON_UNIQUE = 0`,
    [],
    'job schema index check'
  );
  if (indexes.length === 0) {
    try {
      await queryWithRetry(
        'CREATE UNIQUE INDEX idx_jobs_request_id ON jobs (request_id)',
        [],
        'job schema index migration'
      );
    } catch (err) {
      if (err.code !== 'ER_DUP_KEYNAME') throw err;
    }
  }
  schemaReady = true;
}

async function isAvailable() {
  try {
    await queryWithRetry('SELECT 1', [], 'health check', 2);
    return true;
  } catch (err) {
    return false;
  }
}

async function createJob(originalName, requestId) {
  try {
    await ensureSchema();
    const result = await queryWithRetry(
      `INSERT INTO jobs (original_name, status, request_id)
       VALUES (?, 'processing', ?)
       ON DUPLICATE KEY UPDATE
         id = LAST_INSERT_ID(id), original_name = VALUES(original_name)`,
      [originalName, requestId],
      'create job'
    );
    return result.insertId;
  } catch (err) {
    console.warn('Could not create job record (continuing without DB):', formatError(err));
    return null;
  }
}

async function markJobComplete(jobId, { stylesCount, rowsCount, datesCount, outputFile }) {
  if (jobId == null) return;
  try {
    await queryWithRetry(
      `UPDATE jobs
       SET status = 'complete', styles_count = ?, rows_count = ?,
           dates_count = ?, output_file = ?, completed_at = CURRENT_TIMESTAMP
       WHERE id = ?`,
      [stylesCount, rowsCount, datesCount, outputFile, jobId],
      'complete job'
    );
  } catch (err) {
    console.warn('Could not update job record as complete:', formatError(err));
  }
}

async function markJobFailed(jobId, errorMessage) {
  if (jobId == null) return;
  try {
    await queryWithRetry(
      `UPDATE jobs SET status = 'failed', error_message = ?, completed_at = CURRENT_TIMESTAMP WHERE id = ?`,
      [errorMessage, jobId],
      'fail job'
    );
  } catch (err) {
    console.warn('Could not update job record as failed:', formatError(err));
  }
}

async function listRecentJobs(limit = 20) {
  try {
    return await queryWithRetry(
      `SELECT id, original_name, status, styles_count, rows_count, dates_count,
              created_at, completed_at
       FROM jobs
       ORDER BY created_at DESC
       LIMIT ?`,
      [limit],
      'list jobs'
    );
  } catch (err) {
    console.warn('Could not fetch job history (continuing without DB):', formatError(err));
    return [];
  }
}

module.exports = {
  pool,
  isAvailable,
  initializeDatabase: ensureSchema,
  createJob,
  markJobComplete,
  markJobFailed,
  listRecentJobs,
};
