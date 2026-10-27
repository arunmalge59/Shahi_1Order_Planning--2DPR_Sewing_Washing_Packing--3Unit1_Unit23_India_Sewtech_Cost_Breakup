import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { ActivatedRoute, Router, RouterLink } from '@angular/router';
import { SewtechApiService } from '../core/services/sewtech-api.service';

@Component({
  selector: 'app-sewtech-report',
  standalone: true,
  imports: [CommonModule, FormsModule, RouterLink],
  template: `
    <div class="container" *ngIf="report">
      <!-- Top Batch Overview Card -->
      <div class="card">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
          <div>
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 6px;">
              <a routerLink="/cost/sewtech/batches" class="btn secondary" style="padding: 4px 10px; font-size: 12px;">← Batches</a>
              <h2 style="margin: 0;">Sewtech Batch #{{ batchId }} — {{ report.batch.week_label }}</h2>
            </div>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              Date Range: <strong>{{ report.batch.start_date | date:'dd-MMM-yyyy' }}</strong> to <strong>{{ report.batch.end_date | date:'dd-MMM-yyyy' }}</strong> &nbsp;|&nbsp;
              Status: <span class="badge" [ngClass]="{'green': report.batch.status === 'finalized', 'amber': report.batch.status !== 'finalized'}">{{ report.batch.status }}</span>
            </p>
          </div>

          <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center;">
            <button class="btn secondary" (click)="api.downloadMerged(batchId)">
              ⬇ Merged Excel (.xlsx)
            </button>
            <button
              class="btn"
              [disabled]="hasPendingIssues || report.batch.status === 'finalized'"
              (click)="finalizeBatch()"
              [title]="hasPendingIssues ? 'Resolve red action cards below before finalizing' : ''"
            >
              {{ report.batch.status === 'finalized' ? '✓ Batch Finalized' : (hasPendingIssues ? 'Resolve Red Items First' : 'Finalize & Build Summary') }}
            </button>
            <a *ngIf="report.batch.status === 'finalized'" [routerLink]="['/cost/sewtech/final', batchId]" class="btn" style="background:#059669; border-color:#047857;">
              View Final Summary →
            </a>
          </div>
        </div>

        <!-- Metric Summary Tiles -->
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin: 24px 0 12px 0;">
          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #0f172a;">{{ report.totals.total_rows | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Total LINE Rows</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="display: flex; align-items: center; justify-content: space-between;">
              <div style="font-size: 22px; font-weight: 700; color: #0f172a;">{{ report.batch.working_days }}</div>
              <button *ngIf="report.batch.status !== 'finalized'" (click)="toggleEditDays()" class="btn secondary" style="padding: 2px 8px; font-size: 11px;">Edit</button>
            </div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Working Days</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #0e8b94;">{{ report.totals.total_production | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Total Production (pcs)</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #1e6292;">{{ report.totals.total_operators | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Total Operators (sum)</div>
          </div>

          <div style="background: #f8fafc; padding: 14px 16px; border-radius: 10px; border: 1px solid #e2e8f0;">
            <div style="font-size: 22px; font-weight: 700; color: #6a2578;">{{ report.totals.zero_prod_rows | number }}</div>
            <div style="font-size: 12.5px; color: #64748b; margin-top: 2px;">Zero-Prod Kept ({{ report.totals.zero_prod_operators | number }} ops)</div>
          </div>
        </div>

        <!-- Editable Working Days Drawer -->
        <div *ngIf="editingDays" style="background: #fffbeb; border: 1px solid #fde68a; padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; display: flex; gap: 12px; align-items: center;">
          <label style="color: #92400e;"><strong>Edit Working Days:</strong></label>
          <input type="number" [(ngModel)]="customWorkingDays" min="1" max="31" style="width: 80px;" />
          <button class="btn" (click)="saveWorkingDays()">Save & Recalculate</button>
          <button class="btn secondary" (click)="editingDays = false">Cancel</button>
        </div>
      </div>

      <!-- RED ACTION CARDS: Unknown Buyers -->
      <div class="card" *ngIf="report.pending_buyers?.length > 0" style="border-left: 5px solid #ef4444; background: #fffaf9;">
        <h3 style="color: #b91c1c;">⚠ Action Needed: New / Unknown Buyers ({{ report.pending_buyers.length }})</h3>
        <p style="font-size: 13px; color: #64748b;">
          These raw buyers are not in the mapping table. Per business rules, unknown buyers are <strong>never guessed</strong>. Please select or create a Buyer Group.
        </p>

        <table style="background: #ffffff; border: 1px solid #fecaca; border-radius: 8px;">
          <thead>
            <tr style="background: #fef2f2;">
              <th>Raw Buyer in File</th>
              <th>Occurrences</th>
              <th>Sample Dates</th>
              <th>Assign Buyer Group</th>
              <th>Save for Future Weeks?</th>
              <th style="text-align: right;">Action</th>
            </tr>
          </thead>
          <tbody>
            <tr *ngFor="let item of report.pending_buyers">
              <td><span class="badge red">{{ item.raw_buyer }}</span></td>
              <td><strong>{{ item.count }}</strong> rows</td>
              <td>
                <span *ngFor="let d of item.dates" class="badge" style="background:#f1f5f9; margin-right:4px;">{{ d | date:'dd-MMM' }}</span>
              </td>
              <td>
                <div style="display: flex; gap: 8px; align-items: center;">
                  <select [(ngModel)]="buyerGroupSelections[item.raw_buyer]" style="min-width: 140px;">
                    <option value="">-- Choose Group --</option>
                    <option *ngFor="let g of buyerGroups" [value]="g.name">{{ g.name }}</option>
                    <option value="__NEW__">+ Create New Group...</option>
                  </select>
                  <input
                    *ngIf="buyerGroupSelections[item.raw_buyer] === '__NEW__'"
                    type="text"
                    placeholder="New group name"
                    [(ngModel)]="newGroupInputs[item.raw_buyer]"
                    style="width: 140px;"
                  />
                </div>
              </td>
              <td style="text-align: center;">
                <input type="checkbox" [(ngModel)]="saveToMasterFlags[item.raw_buyer]" />
              </td>
              <td style="text-align: right;">
                <button
                  class="btn"
                  style="padding: 4px 12px; font-size: 12px;"
                  [disabled]="!getSelectedGroup(item.raw_buyer)"
                  (click)="resolveBuyer(item.raw_buyer)"
                >
                  Assign & Update
                </button>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- RED ACTION CARDS: Helper (HP) Decision -->
      <div class="card" *ngIf="report.pending_hp_decision" style="border-left: 5px solid #d97706; background: #fffdfa;">
        <h3 style="color: #b45309;">⚠ Action Needed: Helper (HP) Decision Required</h3>
        <p style="font-size: 13px; color: #64748b;">
          Helper (HP) values > 0 were detected in <strong>{{ report.totals.hp_rows }}</strong> rows of this batch.
          Please decide whether to count operators as OP only, or OP + Helpers (HP).
        </p>

        <div style="display: flex; gap: 14px; margin-top: 12px;">
          <button class="btn secondary" (click)="setHpDecision(false)">
            Operators Only (OP)
          </button>
          <button class="btn" style="background: #059669; border-color: #047857;" (click)="setHpDecision(true)">
            Operators + Helpers (OP + HP)
          </button>
        </div>
      </div>

      <!-- Information & Validations Section -->
      <div class="card">
        <h3>Batch Breakdown & Audit Details</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 24px;">
          <!-- Daily production counts -->
          <div>
            <h4 style="margin: 0 0 10px 0; color: #475569; font-size: 13px; text-transform: uppercase;">Daily Breakdown</h4>
            <div style="border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
              <table>
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>LINE Rows</th>
                    <th style="text-align: right;">Production</th>
                    <th style="text-align: right;">Operators</th>
                  </tr>
                </thead>
                <tbody>
                  <tr *ngFor="let day of report.daily_totals">
                    <td><strong>{{ day.work_date | date:'dd-MMM-yyyy' }}</strong></td>
                    <td>{{ day.rows_count }}</td>
                    <td style="text-align: right;">{{ day.day_production | number }}</td>
                    <td style="text-align: right;">{{ day.day_operators | number }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <!-- Change Log & Exclusions -->
          <div>
            <h4 style="margin: 0 0 10px 0; color: #475569; font-size: 13px; text-transform: uppercase;">Audit & Decision Log</h4>
            <div style="max-height: 240px; overflow-y: auto; font-size: 13px; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; background: #fafbfc;">
              <div *ngFor="let log of report.logs" style="margin-bottom: 8px; padding-bottom: 8px; border-bottom: 1px dashed #e2e8f0;">
                <span class="badge" style="background: #e2e8f0; font-size: 11px;">{{ log.category }}</span>
                <span *ngIf="log.category === 'trainee_excluded'"> Trainee rows excluded on {{ log.details?.date }} (count: {{ log.details?.count }}).</span>
                <span *ngIf="log.category === 'zero_prod_kept'"> Retained {{ log.details?.row_count }} zero-production rows ({{ log.details?.operator_count }} operators).</span>
                <span *ngIf="log.category === 'buyer_mapped'"> Assigned '{{ log.details?.raw_buyer }}' → <strong>{{ log.details?.buyer_group }}</strong>.</span>
                <span *ngIf="log.category === 'hp_decision'"> Helper decision: {{ log.details?.description }}.</span>
                <span *ngIf="log.category === 'ignored_tab'"> Ignored sheet: {{ log.details?.tab }} ({{ log.details?.reason }}).</span>
                <span *ngIf="log.category === 'working_days_edited'"> Working days edited to {{ log.details?.confirmed_working_days }}.</span>
              </div>
              <div *ngIf="!report.logs?.length" style="color: #94a3b8;">No change log entries for this batch.</div>
            </div>
          </div>
        </div>
      </div>

      <!-- Merged Rows Preview Table -->
      <div class="card">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
          <div>
            <h3 style="margin: 0 0 4px 0;">Merged Sheet Preview ({{ previewRows.length }} rows)</h3>
            <p style="margin: 0; color: #64748b; font-size: 13px;">Full standardized line-level dataset.</p>
          </div>
          <button class="btn secondary" (click)="api.downloadMerged(batchId)">⬇ Download Full Merged Excel</button>
        </div>

        <div style="overflow-x: auto; max-height: 480px; border: 1px solid #e2e8f0; border-radius: 8px;">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Line</th>
                <th>Unit</th>
                <th style="text-align: right;">Ops</th>
                <th style="text-align: right;">WH</th>
                <th style="text-align: right;">OT Man</th>
                <th style="text-align: right;">OT Hr</th>
                <th>Raw Buyer</th>
                <th>Buyer Group</th>
                <th>Style</th>
                <th>Item no.</th>
                <th style="text-align: right;">Order Qty</th>
                <th style="text-align: right;">SMV</th>
                <th style="text-align: right;">Earn Min</th>
                <th style="text-align: right;">Production</th>
                <th style="text-align: right;">EFFI.</th>
                <th style="text-align: right;">Mc Hrs</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let r of previewRows" [style.background]="r.is_pending ? '#fef2f2' : 'transparent'">
                <td>{{ r.work_date | date:'dd-MMM-yyyy' }}</td>
                <td><strong>{{ r.line }}</strong></td>
                <td>{{ r.unit }}</td>
                <td style="text-align: right;">{{ r.operators_final | number }}</td>
                <td style="text-align: right;">{{ r.reg_hours }}</td>
                <td style="text-align: right;">{{ r.ot_man }}</td>
                <td style="text-align: right;">{{ r.ot_hr }}</td>
                <td>
                  <span *ngIf="r.is_pending" class="badge red">{{ r.raw_buyer || 'Empty' }}</span>
                  <span *ngIf="!r.is_pending">{{ r.raw_buyer }}</span>
                </td>
                <td><strong>{{ r.buyer_group || '—' }}</strong></td>
                <td>{{ r.style }}</td>
                <td>{{ r.item_no }}</td>
                <td style="text-align: right;">{{ r.order_qty | number }}</td>
                <td style="text-align: right;">{{ r.smv }}</td>
                <td style="text-align: right;">{{ r.earn_minutes | number:'1.1-1' }}</td>
                <td style="text-align: right; font-weight: bold; color: #0e8b94;">{{ r.production | number }}</td>
                <td style="text-align: right;">{{ (r.effi * 100) | number:'1.2-2' }}%</td>
                <td style="text-align: right;">{{ r.mc_hrs | number:'1.1-1' }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `,
})
export class SewtechReportComponent implements OnInit {
  batchId!: number;
  report: any = null;
  previewRows: any[] = [];
  buyerGroups: any[] = [];

  buyerGroupSelections: Record<string, string> = {};
  newGroupInputs: Record<string, string> = {};
  saveToMasterFlags: Record<string, boolean> = {};

  editingDays = false;
  customWorkingDays = 6;

  constructor(
    private route: ActivatedRoute,
    private router: Router,
    public api: SewtechApiService
  ) {}

  ngOnInit() {
    this.batchId = Number(this.route.snapshot.paramMap.get('id'));
    this.loadData();
    this.loadBuyerGroups();
  }

  loadData() {
    this.api.getReport(this.batchId).subscribe((res) => {
      this.report = res;
      this.customWorkingDays = res.batch.working_days;

      if (res.pending_buyers) {
        for (const pb of res.pending_buyers) {
          if (this.saveToMasterFlags[pb.raw_buyer] === undefined) {
            this.saveToMasterFlags[pb.raw_buyer] = true;
          }
        }
      }
    });

    this.api.getRows(this.batchId, 300, 0).subscribe((res) => {
      this.previewRows = res.rows || [];
    });
  }

  loadBuyerGroups() {
    this.api.getBuyerGroups().subscribe((groups) => {
      this.buyerGroups = groups || [];
    });
  }

  get hasPendingIssues(): boolean {
    if (!this.report) return false;
    const hasUnresolvedBuyers = (this.report.pending_buyers?.length || 0) > 0;
    const hasPendingHp = !!this.report.pending_hp_decision;
    return hasUnresolvedBuyers || hasPendingHp;
  }

  getSelectedGroup(rawBuyer: string): string {
    const sel = this.buyerGroupSelections[rawBuyer];
    if (sel === '__NEW__') {
      return (this.newGroupInputs[rawBuyer] || '').trim();
    }
    return sel || '';
  }

  resolveBuyer(rawBuyer: string) {
    const targetGroup = this.getSelectedGroup(rawBuyer);
    if (!targetGroup) return;

    this.api.assignBuyer(this.batchId, {
      raw_buyer: rawBuyer,
      buyer_group: targetGroup,
      save_mapping: this.saveToMasterFlags[rawBuyer] !== false,
    }).subscribe(() => {
      this.loadData();
      this.loadBuyerGroups();
    });
  }

  setHpDecision(includeHp: boolean) {
    this.api.decideHp(this.batchId, includeHp).subscribe(() => {
      this.loadData();
    });
  }

  toggleEditDays() {
    this.editingDays = !this.editingDays;
    this.customWorkingDays = this.report?.batch?.working_days || 6;
  }

  saveWorkingDays() {
    if (!this.customWorkingDays || this.customWorkingDays < 1) return;
    this.api.updateWorkingDays(this.batchId, this.customWorkingDays).subscribe(() => {
      this.editingDays = false;
      this.loadData();
    });
  }

  finalizeBatch() {
    if (this.hasPendingIssues) return;
    this.api.finalize(this.batchId).subscribe(() => {
      this.router.navigate(['/cost/sewtech/final', this.batchId]);
    });
  }
}
