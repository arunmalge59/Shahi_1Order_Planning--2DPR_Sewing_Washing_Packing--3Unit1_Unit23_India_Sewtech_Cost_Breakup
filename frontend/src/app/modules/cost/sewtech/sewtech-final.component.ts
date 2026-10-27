import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { ActivatedRoute, RouterLink } from '@angular/router';
import { SewtechApiService } from '../core/services/sewtech-api.service';

@Component({
  selector: 'app-sewtech-final',
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div class="container" *ngIf="batch">
      <div class="card">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 24px;">
          <div>
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 6px;">
              <a [routerLink]="['/cost/sewtech/report', batchId]" class="btn secondary" style="padding: 4px 10px; font-size: 12px;">← Review Grid</a>
              <h2 style="margin: 0;">Sewtech Final Summary — Batch #{{ batchId }}</h2>
            </div>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              Week: <strong>{{ batch.week_label }}</strong> &nbsp;|&nbsp;
              Date Range: <strong>{{ batch.start_date | date:'dd-MMM-yyyy' }}</strong> to <strong>{{ batch.end_date | date:'dd-MMM-yyyy' }}</strong> &nbsp;|&nbsp;
              Working Days Divisor: <strong>{{ batch.working_days }} days</strong>
            </p>
          </div>

          <div style="display: flex; gap: 10px;">
            <button class="btn" (click)="api.downloadFinal(batchId)">⬇ Download Final Summary (.xlsx)</button>
          </div>
        </div>

        <div style="max-width: 860px; margin: 0 auto;">
          <!-- TABLE 1 — SEWTECH BUYER-WISE SUMMARY -->
          <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: 10px; padding: 20px; margin-bottom: 28px; box-shadow: var(--shadow-sm);">
            <div style="font-weight: 700; font-size: 15px; margin-bottom: 12px; color: #0f172a; display: flex; justify-content: space-between; align-items: center;">
              <span>Sewtech Buyer-Wise Final Summary</span>
              <span class="badge teal">{{ batch.working_days }} Working Days</span>
            </div>

            <table style="width: 100%;">
              <thead>
                <tr>
                  <th>Buyer</th>
                  <th style="text-align: right;">{{ batch.week_label || 'Production' }}</th>
                  <th style="text-align: right;">Avg prod/day</th>
                  <th style="text-align: right;">Avg optrs /day</th>
                </tr>
              </thead>
              <tbody>
                <tr *ngFor="let row of summaryRows">
                  <td style="font-weight: 600; color: #0f172a;">{{ row.buyer_group }}</td>
                  <td style="text-align: right;">{{ row.sum_production | number:'1.0-0' }}</td>
                  <td style="text-align: right;">{{ row.avg_prod_per_day | number:'1.0-0' }}</td>
                  <td style="text-align: right;">{{ row.avg_optrs_per_day | number:'1.0-0' }}</td>
                </tr>
              </tbody>
              <tfoot>
                <tr style="font-weight: 700; background: #f8fafc;">
                  <td>Total</td>
                  <td style="text-align: right; color: #0e8b94;">{{ grandTotal.sum_production | number:'1.0-0' }}</td>
                  <td style="text-align: right;">{{ grandTotal.avg_prod_per_day | number:'1.0-0' }}</td>
                  <td style="text-align: right;">{{ grandTotal.avg_optrs_per_day | number:'1.0-0' }}</td>
                </tr>
              </tfoot>
            </table>

            <div style="margin-top: 12px; font-size: 12px; color: #64748b; font-style: italic;">
              * Working days divisor: <strong>{{ batch.working_days }}</strong> days. Metrics generated dynamically from line-level production rows.
            </div>
          </div>

          <!-- TABLE 2 — PRODUCTION & OPERATORS RAW SUMMARY -->
          <div style="background: #ffffff; border: 1px solid #cbd5e1; border-radius: 10px; padding: 20px; box-shadow: var(--shadow-sm);">
            <div style="font-weight: 700; font-size: 14px; margin-bottom: 12px; color: #0f172a;">
              Production & Operators Raw Summary
            </div>

            <table style="width: 100%;">
              <thead>
                <tr>
                  <th>Row Labels</th>
                  <th style="text-align: right;">Sum of Production</th>
                  <th style="text-align: right;">Sum of Operators</th>
                </tr>
              </thead>
              <tbody>
                <tr *ngFor="let row of summaryRows">
                  <td style="font-weight: 500;">{{ row.buyer_group }}</td>
                  <td style="text-align: right;">{{ row.sum_production | number:'1.0-0' }}</td>
                  <td style="text-align: right;">{{ row.sum_operators | number:'1.0-0' }}</td>
                </tr>
              </tbody>
              <tfoot>
                <tr style="font-weight: 700; background: #f8fafc;">
                  <td>Grand Total</td>
                  <td style="text-align: right; color: #0e8b94;">{{ grandTotal.sum_production | number:'1.0-0' }}</td>
                  <td style="text-align: right; color: #1e6292;">{{ grandTotal.sum_operators | number:'1.0-0' }}</td>
                </tr>
              </tfoot>
            </table>

            <div style="margin-top: 12px; font-size: 12px; color: #64748b; font-style: italic;">
              * Represents direct sum aggregations across all parsed line rows for each buyer.
            </div>
          </div>
        </div>
      </div>
    </div>
  `,
})
export class SewtechFinalComponent implements OnInit {
  batchId!: number;
  batch: any = null;
  summaryRows: any[] = [];
  grandTotal = { sum_production: 0, sum_operators: 0, avg_prod_per_day: 0, avg_optrs_per_day: 0 };

  constructor(private route: ActivatedRoute, public api: SewtechApiService) {}

  ngOnInit() {
    this.batchId = Number(this.route.snapshot.paramMap.get('id'));
    this.load();
  }

  load() {
    this.api.getReport(this.batchId).subscribe((res) => {
      this.batch = res.batch;
      if (res.final_summary && res.final_summary.length > 0) {
        this.summaryRows = res.final_summary;
        this.computeGrandTotal();
      }
    });
  }

  computeGrandTotal() {
    const days = Number(this.batch?.working_days) || 6;
    let sumProd = 0;
    let sumOps = 0;

    for (const r of this.summaryRows) {
      sumProd += Number(r.sum_production) || 0;
      sumOps += Number(r.sum_operators) || 0;
    }

    this.grandTotal = {
      sum_production: sumProd,
      sum_operators: sumOps,
      avg_prod_per_day: days > 0 ? sumProd / days : 0,
      avg_optrs_per_day: days > 0 ? sumOps / days : 0,
    };
  }
}
