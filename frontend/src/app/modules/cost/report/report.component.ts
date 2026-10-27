import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { ActivatedRoute, RouterLink } from '@angular/router';
import { ApiService } from '../core/services/api.service';

@Component({
  selector: 'app-report',
  standalone: true,
  imports: [CommonModule, FormsModule, RouterLink],
  template: `
    <div class="container" *ngIf="report">
      <!-- Top Batch Overview Card -->
      <div class="card">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
          <div>
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 6px;">
              <a routerLink="/cost/india/batches" class="btn secondary" style="padding: 4px 10px; font-size: 12px;">← Batches</a>
              <h2 style="margin: 0;">India Batch #{{ batchId }} — {{ report.batch.week_label }}</h2>
            </div>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              Status: <span class="badge" [ngClass]="{'green': report.batch.status === 'finalized', 'amber': report.batch.status !== 'finalized'}">{{ report.batch.status }}</span> &nbsp;|&nbsp;
              Working days used: <strong>{{ report.batch.working_days }} days</strong>
            </p>
          </div>

          <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">
            <button class="btn secondary" (click)="api.downloadModification(batchId)">
              ⬇ Modification Sheet (.xlsx)
            </button>
            <button
              class="btn"
              [disabled]="report.new_brand.length > 0 || report.batch.status === 'finalized'"
              (click)="finalize()"
            >
              {{ report.batch.status === 'finalized' ? '✓ Batch Finalized' : (report.new_brand.length > 0 ? 'Resolve New Brands Below' : 'Finalize & Build Summary') }}
            </button>
          </div>
        </div>

        <!-- Metric Tiles -->
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin: 24px 0 8px 0;">
          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #0e8b94;">{{ report.totals.kept_rows | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Rows Kept</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #64748b;">{{ report.totals.deleted_rows | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Deleted (Qty ≤ 1)</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700;" [style.color]="report.totals.new_brand_rows > 0 ? '#ef4444' : '#059669'">
              {{ report.totals.new_brand_rows | number }}
            </div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">New Brands Flagged</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #6a2578;">{{ report.totals.sam_fixed_rows | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">SAM / Eff = 0 Fixed</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #1e6292;">{{ report.forward_fill_count | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Names Forward-Filled</div>
          </div>
        </div>
      </div>

      <!-- Action Card: New Brands -->
      <div class="card" *ngIf="report.new_brand.length" style="border-left: 5px solid #ef4444; background: #fffaf9;">
        <h3 style="color: #b91c1c;">⚠ New Brands — Not in Mapping Master, Action Required</h3>
        <p style="font-size: 13px; color: #64748b;">
          These brands were not recognized in the system. Please assign each to its designated Buyer Group.
        </p>
        <table style="background: #ffffff; border: 1px solid #fecaca; border-radius: 8px;">
          <thead>
            <tr style="background: #fef2f2;">
              <th>Excel Row No.</th>
              <th>Brand</th>
              <th>Item</th>
              <th>Assign to Buyer Group</th>
              <th>Save to Master?</th>
              <th style="text-align: right;">Action</th>
            </tr>
          </thead>
          <tbody>
            <tr *ngFor="let entry of report.new_brand">
              <td><strong>#{{ entry.raw_row_no || entry.details?.raw_row_no || '-' }}</strong></td>
              <td><span class="badge red">{{ entry.brand }}</span></td>
              <td>{{ entry.item }}</td>
              <td><input type="text" [(ngModel)]="assignInputs[entry.brand]" placeholder="e.g. Kohls" style="width: 100%; max-width: 180px;" /></td>
              <td style="text-align: center;"><input type="checkbox" [(ngModel)]="saveToMaster[entry.brand]" checked /></td>
              <td style="text-align: right;"><button class="btn secondary" (click)="assign(entry.brand, undefined)">Assign</button></td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Split Brands -->
      <div class="card" *ngIf="report.buyer_split_needed.length">
        <h3>Split-Brand Items Needing Confirmation (e.g. ADIDAS / H&M Denim vs Non-Denim)</h3>
        <table>
          <thead>
            <tr>
              <th>Excel Row No.</th>
              <th>Brand</th>
              <th>Item</th>
              <th>Message</th>
              <th>Assign Group</th>
              <th style="text-align: right;">Action</th>
            </tr>
          </thead>
          <tbody>
            <tr *ngFor="let entry of report.buyer_split_needed">
              <td><strong>#{{ entry.raw_row_no || entry.details?.raw_row_no || '-' }}</strong></td>
              <td>{{ entry.brand }}</td>
              <td>{{ entry.item }}</td>
              <td>{{ entry.details.message }}</td>
              <td><input type="text" [(ngModel)]="assignInputs[entry.brand + '||' + entry.item]" placeholder="e.g. Adidas denim" /></td>
              <td style="text-align: right;"><button class="btn secondary" (click)="assign(entry.brand, entry.item)">Assign</button></td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Rows Deleted -->
      <div class="card" *ngIf="report.row_deleted.length">
        <h3>Rows Excluded (Production Qty ≤ 1) — {{ report.row_deleted.length }} records</h3>
        <div style="max-height: 280px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 8px;">
          <table>
            <thead>
              <tr>
                <th>Excel Row No.</th>
                <th>Brand</th>
                <th>Item</th>
                <th>Production Qty</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let entry of report.row_deleted.slice(0, 100)">
                <td><strong>#{{ entry.raw_row_no || entry.details?.raw_row_no || '-' }}</strong></td>
                <td>{{ entry.brand }}</td>
                <td>{{ entry.item }}</td>
                <td>{{ entry.details.production_qty }}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p *ngIf="report.row_deleted.length > 100" style="margin-top: 8px; font-size: 12px; color: #64748b;">Showing first 100 of {{ report.row_deleted.length }} rows.</p>
      </div>

      <!-- SAM Fixed -->
      <div class="card" *ngIf="report.sam_eff_fixed.length">
        <h3>SAM = 0 / Efficiency = 0 Rows Handled</h3>
        <table>
          <thead>
            <tr>
              <th>Excel Row No.</th>
              <th>Brand</th>
              <th>Item</th>
              <th>Action Applied / Missing Master</th>
            </tr>
          </thead>
          <tbody>
            <tr *ngFor="let entry of report.sam_eff_fixed">
              <td><strong>#{{ entry.raw_row_no || entry.details?.raw_row_no || '-' }}</strong></td>
              <td>{{ entry.brand }}</td>
              <td>{{ entry.item }}</td>
              <td>
                <span *ngIf="entry.details?.needs_sam_master" class="badge red" style="margin-right:8px;">Action Required</span>
                <span *ngIf="entry.details?.needs_sam_master">SAM master entry missing. Manually verify and enter the standard SAM value in Master Data.</span>
                <span *ngIf="!entry.details?.needs_sam_master">
                  Std SAM/pc {{ entry.details?.standard_sam_per_piece }} × Qty {{ entry.details?.production_qty }}
                  = SAM {{ entry.details?.new_sam }}, Efficiency set to {{ entry.details?.new_efficiency }}
                </span>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Final Pivot Summary -->
      <div class="card" *ngIf="finalPivot">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; margin-bottom: 20px; gap: 12px;">
          <div>
            <h3 style="margin: 0 0 4px 0;">Final Buyer-Wise Production Summary</h3>
            <p style="margin: 0; color: #64748b; font-size: 13px;">Week: <strong>{{ report?.batch?.week_label }}</strong> | Working Days Divisor: <strong>{{ report?.batch?.working_days }} days</strong></p>
          </div>
          <button class="btn" (click)="api.downloadFinal(batchId)">⬇ Download Final Summary (.xlsx)</button>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 24px; align-items: flex-start; margin-bottom: 20px;">
          <!-- Left Table (Pivot) -->
          <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 16px;">
            <div style="font-weight: 700; font-size: 13.5px; margin-bottom: 12px; color: #0f172a;">Production Qty & Mcs Pivot</div>
            <table style="width: 100%;">
              <thead>
                <tr>
                  <th>Buyer Group</th>
                  <th style="text-align: right;">Sum of Production Qty</th>
                  <th style="text-align: right;">Sum of Mcs</th>
                </tr>
              </thead>
              <tbody>
                <tr *ngFor="let g of alphabetizedFinalPivot">
                  <td style="font-weight: 500;">{{ g.buyer_group }}</td>
                  <td style="text-align: right;">{{ roundVal(g.sum_production_qty) | number }}</td>
                  <td style="text-align: right;">{{ roundVal(g.sum_mcs) | number }}</td>
                </tr>
              </tbody>
              <tfoot>
                <tr style="font-weight: 700; background: #f8fafc;">
                  <td>Grand Total</td>
                  <td style="text-align: right;">{{ totalProductionQty | number }}</td>
                  <td style="text-align: right;">{{ totalMcs | number }}</td>
                </tr>
              </tfoot>
            </table>
          </div>

          <!-- Right Table (Daily Averages) -->
          <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; padding: 16px;">
            <div style="font-weight: 700; font-size: 13.5px; margin-bottom: 12px; color: #0f172a;">Daily Average Production & Operators</div>
            <table style="width: 100%;">
              <thead>
                <tr>
                  <th>Buyer</th>
                  <th style="text-align: right;">Production Qty</th>
                  <th style="text-align: right;">Avg prod/day</th>
                  <th style="text-align: right;">Avg optrs /day</th>
                </tr>
              </thead>
              <tbody>
                <tr *ngFor="let g of sortedFinalPivot">
                  <td style="font-weight: 500;">{{ g.buyer_group }}</td>
                  <td style="text-align: right;">{{ roundVal(g.sum_production_qty) | number }}</td>
                  <td style="text-align: right;">{{ getAvgProd(g) | number }}</td>
                  <td style="text-align: right;">{{ getAvgOptrs(g) | number }}</td>
                </tr>
              </tbody>
              <tfoot>
                <tr style="font-weight: 700; background: #f8fafc;">
                  <td>Total</td>
                  <td style="text-align: right;">{{ totalProductionQty | number }}</td>
                  <td style="text-align: right;">{{ totalAvgProdPerDay | number }}</td>
                  <td style="text-align: right;">{{ totalAvgOptrsPerDay | number }}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      </div>
    </div>
  `,
})
export class ReportComponent implements OnInit {
  batchId!: number;
  report: any;
  finalPivot: any[] | null = null;
  assignInputs: Record<string, string> = {};
  saveToMaster: Record<string, boolean> = {};

  constructor(private route: ActivatedRoute, public api: ApiService) {}

  ngOnInit() {
    this.batchId = Number(this.route.snapshot.paramMap.get('id'));
    this.load();
  }

  load() {
    this.api.getReport(this.batchId).subscribe((r) => {
      this.report = r;
      if (r.finalPivot) {
        this.finalPivot = r.finalPivot;
      }
    });
  }

  get alphabetizedFinalPivot(): any[] {
    if (!this.finalPivot) return [];
    return [...this.finalPivot].sort((a, b) => (a.buyer_group || '').localeCompare(b.buyer_group || ''));
  }

  get sortedFinalPivot(): any[] {
    if (!this.finalPivot) return [];
    return [...this.finalPivot].sort((a, b) => (Number(b.sum_production_qty) || 0) - (Number(a.sum_production_qty) || 0));
  }

  roundVal(val: any): number {
    return Math.round(Number(val) || 0);
  }

  getAvgProd(g: any): number {
    if (g.avg_prod_per_day != null && g.avg_prod_per_day !== '') {
      return Math.round(Number(g.avg_prod_per_day));
    }
    const days = Number(this.report?.batch?.working_days) || 5;
    return Math.round((Number(g.sum_production_qty) || 0) / days);
  }

  getAvgOptrs(g: any): number {
    const val = Number(g.avg_optrs_per_day);
    return Math.round(Number.isFinite(val) ? val : 0);
  }

  get totalProductionQty(): number {
    const rawSum = this.sortedFinalPivot.reduce(
      (sum, g) => sum + (Number(g.sum_production_qty) || 0),
      0
    );
    return Math.round(rawSum);
  }

  get totalMcs(): number {
    const rawSum = this.sortedFinalPivot.reduce(
      (sum, g) => sum + (Number(g.sum_mcs) || 0),
      0
    );
    return Math.round(rawSum);
  }

  get totalAvgProdPerDay(): number {
    const rawSum = this.sortedFinalPivot.reduce((sum, g) => {
      const val =
        g.avg_prod_per_day != null && g.avg_prod_per_day !== ''
          ? Number(g.avg_prod_per_day)
          : (Number(g.sum_production_qty) || 0) /
            (Number(this.report?.batch?.working_days) || 5);

      return sum + (Number.isFinite(val) ? val : 0);
    }, 0);

    return Math.round(rawSum);
  }

  get totalAvgOptrsPerDay(): number {
    const rawSum = this.sortedFinalPivot.reduce((sum, g) => {
      const val = Number(g.avg_optrs_per_day);
      return sum + (Number.isFinite(val) ? val : 0);
    }, 0);

    return Math.round(rawSum);
  }

  assign(brand: string, item?: string) {
    const key = item ? `${brand}||${item}` : brand;
    const buyerGroup = this.assignInputs[key];
    if (!buyerGroup) return;
    this.api
      .assignBuyer(this.batchId, { brand, item, buyer_group: buyerGroup, save_to_master: this.saveToMaster[brand] ?? true })
      .subscribe(() => this.load());
  }

  finalize() {
    this.api.finalize(this.batchId).subscribe((res) => (this.finalPivot = res.pivot));
  }
}
