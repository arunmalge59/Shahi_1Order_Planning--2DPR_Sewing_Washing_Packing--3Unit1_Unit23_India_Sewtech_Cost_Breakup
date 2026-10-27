const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '../../.env') });

const express = require('express');

const productionApp = require('../modules/production/src/server');
const costApp = require('../modules/cost/src/server');
const planningApp = require('../modules/planning/src/server');
const productionDb = require('../modules/production/src/db');

const app = express();
const port = Number(process.env.PORT || 4000);

app.get('/api/health', async (req, res) => {
  res.json({
    ok: true,
    database: await productionDb.isAvailable(),
    modules: ['production', 'cost', 'planning']
  });
});

app.use('/production', productionApp);
app.use('/cost', costApp);
app.use('/planning', planningApp);

const frontendDistPath = path.resolve(__dirname, '../../frontend/dist/shahi-unified-operations/browser');
app.use(express.static(frontendDistPath));
app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api/') || req.path.startsWith('/production/api/') ||
      req.path.startsWith('/cost/api/') || req.path.startsWith('/planning/api/')) {
    return next();
  }
  res.sendFile(path.join(frontendDistPath, 'index.html'), (error) => {
    if (error) next();
  });
});

let server;
if (require.main === module) {
  server = app.listen(port, '0.0.0.0', () => {
    console.log(`Shahi Unified Operations API listening on port ${port}`);
  });
}

module.exports = { app, get server() { return server; } };
