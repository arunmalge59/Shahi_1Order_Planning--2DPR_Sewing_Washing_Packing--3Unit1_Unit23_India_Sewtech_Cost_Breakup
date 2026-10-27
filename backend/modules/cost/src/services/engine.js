/**
 * Core engine that replicates the manual weekly process your manager does in Excel:
 *
 *   1. Forward-fill Sales Person / Item Responsible down blank rows.
 *   2. Resolve Buyer GROUP from Brand (+ Item, for split brands like ADIDAS/H&M) using
 *      the brand_buyer_map / item_buyer_override tables — NOT the raw "Buyer" column.
 *      Unknown brands are flagged (is_new_brand = true) instead of guessed.
 *   3. Delete rows where Production Qty <= 1 (nothing meaningfully in production).
 *   4. Fix rows where SAM = 0 and Actual Efficiency = 0 (new style, no master data yet):
 *        sam := production_qty * standard_sam_per_piece   (from sam_master; if missing, flagged)
 *        actual_efficiency := DEFAULT_EFFICIENCY_FOR_ZERO (default 20)
 *   5. Compute Mcs = ((sam / actual_efficiency) * 100) / (480 * working_days)
 *   6. Actual Cost/SAM is left untouched — it already comes from the ERP export as
 *      total_cost / sam and is carried straight through.
 *
 * Every decision made here is written to change_log so the "what changed" report
 * can be shown to the user right after upload, before anything is finalized.
 */

const DEFAULT_EFFICIENCY_FOR_ZERO = Number(process.env.DEFAULT_EFFICIENCY_FOR_ZERO || 20);

function forwardFill(rows) {
  let lastSalesPerson = null;
  let lastItemResponsible = null;
  const log = [];

  for (const row of rows) {
    if (row.sales_person && String(row.sales_person).trim() !== '') {
      lastSalesPerson = row.sales_person;
    } else {
      if (lastSalesPerson) {
        row.sales_person = lastSalesPerson;
        log.push({ category: 'forward_fill', item: row.item, raw_row_no: row.raw_row_no, details: { raw_row_no: row.raw_row_no, field: 'sales_person', filled_with: lastSalesPerson } });
      }
    }

    if (row.item_responsible && String(row.item_responsible).trim() !== '') {
      lastItemResponsible = row.item_responsible;
    } else {
      if (lastItemResponsible) {
        row.item_responsible = lastItemResponsible;
        log.push({ category: 'forward_fill', item: row.item, raw_row_no: row.raw_row_no, details: { raw_row_no: row.raw_row_no, field: 'item_responsible', filled_with: lastItemResponsible } });
      }
    }
  }
  return log;
}

/**
 * brandMap: Map<brand, { buyer_group, is_denim_split }>
 * overrideMap: Map<`${brand}||${item}`, buyer_group>
 */
function resolveBuyers(rows, brandMap, overrideMap) {
  const log = [];

  for (const row of rows) {
    const brand = (row.brand || '').trim();
    const override = overrideMap.get(`${brand}||${row.item}`);

    if (override) {
      row.resolved_buyer = override;
      row.is_new_brand = false;
      continue;
    }

    const mapped = brandMap.get(brand);
    if (mapped) {
      row.resolved_buyer = mapped.buyer_group;
      row.is_new_brand = false;
      if (mapped.is_denim_split) {
        // Known split-brand (e.g. ADIDAS, H&M) but this specific item has no override yet.
        // Falls back to the brand's default group but still needs a human decision.
        log.push({
          category: 'buyer_split_needed',
          item: row.item,
          brand,
          raw_row_no: row.raw_row_no,
          details: {
            raw_row_no: row.raw_row_no,
            message: `${brand} splits into multiple buyer groups; item "${row.item}" has no saved override, defaulted to "${mapped.buyer_group}". Please confirm.`,
          },
        });
      }
    } else {
      // Brand never seen before -> do NOT guess, flag it red for manual review.
      row.resolved_buyer = row.raw_buyer || null;
      row.is_new_brand = true;
      log.push({
        category: 'new_brand',
        item: row.item,
        brand,
        raw_row_no: row.raw_row_no,
        details: {
          raw_row_no: row.raw_row_no,
          message: `Brand "${brand}" not found in brand_buyer_map. Row kept as-is and flagged for manual mapping.`,
        },
      });
    }
  }
  return log;
}

function filterAndDeleteLowQty(rows) {
  const kept = [];
  const log = [];
  for (const row of rows) {
    const qty = Number(row.production_qty);
    if (isNaN(qty) || qty <= 1) {
      row.was_deleted = true;
      row.delete_reason = `Production Qty (${row.production_qty}) <= 1`;
      log.push({
        category: 'row_deleted',
        item: row.item,
        brand: row.brand,
        raw_row_no: row.raw_row_no,
        details: { raw_row_no: row.raw_row_no, production_qty: row.production_qty },
      });
      kept.push(row); // kept in the audit table with was_deleted=true, excluded from final calcs
    } else {
      row.was_deleted = false;
      kept.push(row);
    }
  }
  return { rows: kept, log };
}

/**
 * samMasterMap: Map<item, standard_sam_per_piece>
 */
function fixZeroSamAndEfficiency(rows, samMasterMap) {
  const log = [];
  for (const row of rows) {
    if (row.was_deleted) continue;
    const sam = Number(row.sam);
    const eff = Number(row.actual_efficiency);

    if ((sam === 0 || isNaN(sam)) && (eff === 0 || isNaN(eff))) {
      const stdSam = samMasterMap.get(row.item);
      if (stdSam != null) {
        const newSam = Number(row.production_qty) * Number(stdSam);
        row.sam = newSam;
        row.actual_efficiency = DEFAULT_EFFICIENCY_FOR_ZERO;
        row.was_sam_zero_fixed = true;
        log.push({
          category: 'sam_eff_fixed',
          item: row.item,
          brand: row.brand,
          raw_row_no: row.raw_row_no,
          details: {
            raw_row_no: row.raw_row_no,
            standard_sam_per_piece: stdSam,
            production_qty: row.production_qty,
            new_sam: newSam,
            new_efficiency: DEFAULT_EFFICIENCY_FOR_ZERO,
            message: `Std SAM/pc ${stdSam} × Qty ${row.production_qty} = SAM ${newSam}, Efficiency set to ${DEFAULT_EFFICIENCY_FOR_ZERO}`,
          },
        });
      } else {
        row.was_sam_zero_fixed = false;
        log.push({
          category: 'sam_eff_fixed',
          item: row.item,
          brand: row.brand,
          raw_row_no: row.raw_row_no,
          details: {
            raw_row_no: row.raw_row_no,
            needs_sam_master: true,
            message: 'SAM master entry missing. Manually verify and enter the standard SAM value.',
          },
        });
      }
    }
  }
  return log;
}

function computeMcs(rows, workingDays) {
  for (const row of rows) {
    if (row.was_deleted) continue;
    const sam = Number(row.sam);
    const eff = Number(row.actual_efficiency);
    if (!eff || isNaN(sam)) {
      row.mcs = null;
      continue;
    }
    row.mcs = ((sam / eff) * 100) / (480 * workingDays);
  }
}

/**
 * Full pipeline: raw rows in -> cleaned rows + change log out.
 * Does not touch Actual Cost/SAM — that column is carried straight through from the ERP export.
 */
function runPipeline({ rawRows, brandMap, overrideMap, samMasterMap, workingDays }) {
  const rows = rawRows.map((r) => ({ ...r }));
  const changeLog = [];

  changeLog.push(...forwardFill(rows));
  changeLog.push(...resolveBuyers(rows, brandMap, overrideMap));

  const { rows: filteredRows, log: deleteLog } = filterAndDeleteLowQty(rows);
  changeLog.push(...deleteLog);

  changeLog.push(...fixZeroSamAndEfficiency(filteredRows, samMasterMap));
  computeMcs(filteredRows, workingDays);

  return { rows: filteredRows, changeLog };
}

/**
 * Pivot by resolved_buyer, matching the Final_INDIAfile.xlsx sheet.
 */
function pivotByBuyer(rows, workingDays) {
  const groups = new Map();
  for (const row of rows) {
    if (row.was_deleted) continue;
    const buyer = row.resolved_buyer || 'UNMAPPED';
    if (!groups.has(buyer)) {
      groups.set(buyer, { buyer_group: buyer, sum_production_qty: 0, sum_mcs: 0 });
    }
    const g = groups.get(buyer);
    g.sum_production_qty += Number(row.production_qty) || 0;
    g.sum_mcs += Number(row.mcs) || 0;
  }

  return Array.from(groups.values()).map((g) => ({
    ...g,
    avg_prod_per_day: g.sum_production_qty / workingDays,
    avg_optrs_per_day: g.sum_mcs,
  }));
}

module.exports = {
  DEFAULT_EFFICIENCY_FOR_ZERO,
  forwardFill,
  resolveBuyers,
  filterAndDeleteLowQty,
  fixZeroSamAndEfficiency,
  computeMcs,
  runPipeline,
  pivotByBuyer,
};
