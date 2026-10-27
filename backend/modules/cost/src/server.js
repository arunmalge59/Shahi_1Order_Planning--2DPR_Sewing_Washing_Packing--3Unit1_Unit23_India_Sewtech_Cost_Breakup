require('dotenv').config();
require('express-async-errors');

const express = require('express');
const cors = require('cors');

const batchesRouter = require('./routes/batches');
const masterDataRouter = require('./routes/masterData');
const sewtechRouter = require('./routes/sewtech');

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

app.get('/api/health', (req, res) => res.json({ ok: true }));

app.use('/api/batches', batchesRouter);
app.use('/api/master', masterDataRouter);
app.use('/api/sewtech', sewtechRouter);

// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: err.message || 'Internal server error' });
});

module.exports = app;
