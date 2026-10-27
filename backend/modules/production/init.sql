-- Line Plan Generator: job history table
CREATE TABLE IF NOT EXISTS jobs (
  id            INT NOT NULL AUTO_INCREMENT,
  original_name TEXT NOT NULL,
  status        VARCHAR(20) NOT NULL DEFAULT 'processing',
  styles_count  INT,
  rows_count    INT,
  dates_count   INT,
  output_file   TEXT,
  error_message TEXT,
  request_id    VARCHAR(255),
  created_at    TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  completed_at  TIMESTAMP NULL DEFAULT NULL,
  PRIMARY KEY (id),
  KEY idx_jobs_created_at (created_at DESC),
  UNIQUE KEY idx_jobs_request_id (request_id)
);
