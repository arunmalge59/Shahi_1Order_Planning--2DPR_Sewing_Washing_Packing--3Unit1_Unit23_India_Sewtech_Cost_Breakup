const assert = require('node:assert/strict');
const { after, before, test } = require('node:test');
const XLSX = require('xlsx');
const { app } = require('../src/server');
const { compareSnapshots } = require('../modules/planning/src/compare');
const { runPipeline } = require('../modules/cost/src/services/engine');
const costPool = require('../modules/cost/src/db/pool');
const productionDb = require('../modules/production/src/db');
const planningDb = require('../modules/planning/src/db');

let server;
let baseUrl;

before(async () => {
  server = app.listen(0);
  await new Promise((resolve) => server.once('listening', resolve));
  baseUrl = `http://127.0.0.1:${server.address().port}`;
});

after(async () => {
  await new Promise((resolve, reject) => {
    server.close((error) => error ? reject(error) : resolve());
  });
  await Promise.all([
    costPool.end(),
    productionDb.pool.end(),
    planningDb.pool.end()
  ]);
});

test('unified API reports all three modules', async () => {
  const response = await fetch(`${baseUrl}/api/health`);
  assert.equal(response.status, 200);
  const body = await response.json();
  assert.deepEqual(body.modules, ['production', 'cost', 'planning']);
  assert.equal(typeof body.database, 'boolean');
});

test('each module is reachable through its unified API prefix', async () => {
  for (const [module, endpoint] of [
    ['production', '/api/health'],
    ['cost', '/api/health'],
    ['planning', '/api/health']
  ]) {
    const response = await fetch(`${baseUrl}/${module}${endpoint}`);
    assert.equal(response.status, 200, `${module} health endpoint`);
  }
});

test('planning comparison preserves the quantity-change calculation', () => {
  const result = compareSnapshots(
    [{ style: 'ST-100', description: 'JACKET', qty: 100, pcd: '01-02-2026', exFactory: '03-02-2026', itemNumber: 'ITEM-1' }],
    [{ style: 'ST-100', description: 'JACKET', qty: 150, pcd: '01-02-2026', exFactory: '03-02-2026', itemNumber: 'ITEM-1' }]
  );

  assert.equal(result.summary.qtyChanged, 1);
  assert.equal(result.styles[0].qtyDelta, 50);
  assert.equal(result.styles[0].status, 'HIGH');
});

test('cost pipeline preserves buyer resolution and MCS calculation', () => {
  const result = runPipeline({
    rawRows: [{
      item: 'ITEM-1',
      brand: 'ACME',
      production_qty: 20,
      sam: 80,
      actual_efficiency: 50,
      actual_cost_sam: 12
    }],
    brandMap: new Map([['ACME', { buyer_group: 'Group A', is_denim_split: false }]]),
    overrideMap: new Map(),
    samMasterMap: new Map(),
    workingDays: 20
  });

  assert.equal(result.rows[0].resolved_buyer, 'Group A');
  assert.equal(result.rows[0].actual_cost_sam, 12);
  assert.ok(Math.abs(result.rows[0].mcs - (80 / 50 * 100 / (480 * 20))) < 1e-12);
});

test('planning API imports an Excel workbook and exports the generated order CSV', async () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Style', 'SEASON', 'Qty', 'PCD', 'Revised PCD', 'EX FACTORY', 'Buy date', 'Description'],
    ['ST-100', 'FA', 2000, '01-02-2026', '02-02-2026', '03-02-2026', '04-02-2026', 'JACKET']
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary sheet');
  const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const form = new FormData();
  form.append('file', new Blob([excelBuffer]), 'sample.xlsx');
  form.append('weekCode', 'W42');

  const response = await fetch(`${baseUrl}/planning/api/order-file/generate`, {
    method: 'POST',
    body: form
  });
  const csv = await response.text();

  assert.equal(response.status, 200);
  assert.match(csv, /^O\.CODE,::,/);
  assert.match(csv, /ST-100::FA'27::20260203::W42/);
  assert.match(csv, /,2060,2000,/);
});
