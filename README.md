# Shahi Unified Operations

One Angular application, one Express process, and one MySQL database for the three existing Shahi workflows. Original project folders remain unchanged as backups.

## Modules

- **Daily DPR**: line plan generation, sewing actuals, wash/pack actuals, order summary, Unit 23 workflows, exclusions, job history, and workbook downloads.
- **India Cost Breakup**: India costing and Sewtech uploads, batches, review/finalization, exports, buyer mappings, and master data.
- **Planning and Order**: Nike snapshot and category comparisons, Nike/Adidas order files, H&M denim and non-denim order-file generation, and chart variations.

The existing workbook processing logic is retained within its module. The shared Angular navigation has exactly three expandable main menus:

- **Daily DPR**: Unit 1 and Unit 23 plan and actuals workflows.
- **India Cost Breakup**: India Factory and Sewtech uploads, batches, reports, and master data.
- **Planning and Order**: the existing Nike, Adidas, and H&M comparison and order-file workflows.

All existing routes and APIs remain available; the navigation does not use iframes.

## Requirements

- Node.js 20 or newer and npm
- MySQL Server 8 (MySQL Workbench can be used to manage the server)
- Python 3 for Production Operations (`python -m pip install openpyxl`)

## Setup

1. From this directory, install dependencies:

   ```powershell
   npm install
   ```

2. Copy `.env.example` to `.env` and set `DB_HOST`, `DB_PORT`, `DB_USER`, and `DB_PASSWORD`. The database is named `shahi_unified_operations_db` by default.

3. Ensure the configured MySQL user can create the database and tables, then run the safe, additive schema setup:

   ```powershell
   npm run db:migrate
   ```

   The migration uses `CREATE DATABASE IF NOT EXISTS` and `CREATE TABLE IF NOT EXISTS`; it does not drop tables or delete rows.

4. Start the unified API and Angular development server:

   ```powershell
   npm run dev
   ```

   Open `http://localhost:4200`. The API listens on `http://localhost:4000`.

## Validation

```powershell
npm run build
npm run test:backend
```

The backend tests include the original synthetic Sewtech workbook regression scenario and India costing calculations. Live database connectivity and processing user-supplied workbooks require a running MySQL Server, Python with the production module's required packages, and representative input workbooks.

## Unified layout

```text
Shahi_Unified_Operations/
  backend/
    modules/
      production/   # Existing production API, Python workbook logic, schema
      cost/         # Existing India and Sewtech APIs, calculations, schemas
      planning/     # Existing order and comparison API logic, schema
    src/
      db/migrate.js
      server.js     # Single Express listener
    test/
  frontend/
    src/
      app/
        app.component.ts
        app.routes.ts
        modules/
          production/
          cost/
          planning/
      environments/
  .env.example
  package.json
```
