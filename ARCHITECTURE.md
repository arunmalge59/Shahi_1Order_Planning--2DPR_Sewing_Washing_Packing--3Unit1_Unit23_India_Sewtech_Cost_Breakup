# Architecture and integration map

## Runtime layout

- `frontend/` is a single Angular 18 standalone application with a shared responsive Tailwind layout and exactly three top-level navigation menus.
- `backend/src/server.js` is the only Express listener. It mounts the existing API applications beneath `/production`, `/cost`, and `/planning`.
- All three module database clients use the same `.env` MySQL settings and default to `shahi_unified_operations_db`. Their table families remain separate to preserve each module's data model.
- `backend/src/db/migrate.js` creates the shared database if needed, then ensures all module schemas using additive `CREATE ... IF NOT EXISTS` operations.

## Main navigation and workflows

| Main menu | Routes | Preserved workflows |
|---|---|---|
| Daily DPR | `/production` | Unit 1 and Unit 23 plan generation and actuals uploads |
| India Cost Breakup | `/cost/india/*`, `/cost/sewtech/*` | India and Sewtech costing uploads, batch review, buyer assignments, master data, finalization and Excel exports |
| Planning and Order | `/planning/*` | Nike snapshot and category comparisons, Nike/Adidas order files, and H&M denim/non-denim workflows |

The Daily DPR menu groups its Unit 1 and Unit 23 workflows separately. The India Cost Breakup menu groups India Factory and Sewtech workflows. The Planning and Order menu keeps all existing customer comparison and order-file workflows. All route URLs and module API namespaces remain unchanged.

## API mount points

Each existing module API is exposed under its namespace, for example:

- Production: `/production/api/*`
- Cost: `/cost/api/*`
- Planning: `/planning/api/*`
- Unified health and database status: `/api/health`

The original files are copied into `backend/modules` and `frontend/src/app/modules`; source project directories are not modified. No iframe or mock workflow is used.

## Shared database tables

- Production Operations: `jobs`
- Cost Analysis / India: `brand_buyer_map`, `item_buyer_override`, `sam_master`, `batches`, `batch_rows`, `change_log`, `final_summary`
- Cost Analysis / Sewtech: `sewtech_*` tables
- Order & Planning: `comparison_runs`, `comparison_style_changes`

Names remain module-owned and distinct. Migration scripts only create missing tables and do not delete records.

## Costing regression coverage

`backend/test/cost-regression.test.js` runs the original synthetic Sewtech workbook scenario against the integrated costing services. It checks date scanning and range filtering, all 210 expected rows, trainee exclusion, zero-production retention, buyer mapping and summaries, unknown-buyer flags, Excel exports, and the India processing pipeline without requiring a live database.
