import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { RouterLink } from '@angular/router';
import { SewtechApiService } from '../core/services/sewtech-api.service';

@Component({
  selector: 'app-sewtech-batches',
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div class="container">
      <div class="card">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
          <div>
            <h2 style="margin: 0 0 4px 0;">Sewtech Weekly Batches</h2>
            <p style="margin: 0; color: #64748b; font-size: 13px;">Manage, review, and finalize Sewtech factory weekly upload batches.</p>
          </div>
          <a routerLink="/cost/sewtech/upload" class="btn">+ New Weekly Upload</a>
        </div>

        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>Week / Dates</th>
              <th>Working Days</th>
              <th>Status</th>
              <th>Raw File</th>
              <th>Uploaded At</th>
              <th style="text-align: right;">Actions</th>
            </tr>
          </thead>
          <tbody>
            <tr *ngFor="let b of batches">
              <td><strong>#{{ b.id }}</strong></td>
              <td>
                <div style="font-weight: 600; color: var(--text-primary);">{{ b.week_label }}</div>
                <div style="font-size: 11px; color: #64748b; margin-top: 2px;">
                  {{ b.start_date | date:'dd-MMM' }} to {{ b.end_date | date:'dd-MMM-yyyy' }}
                </div>
              </td>
              <td>{{ b.working_days }} days</td>
              <td>
                <span
                  class="badge"
                  [ngClass]="{
                    'green': b.status === 'finalized',
                    'amber': b.status === 'raw_uploaded' || b.status === 'processed',
                    'purple': b.status === 'modified_uploaded'
                  }"
                >
                  {{ b.status }}
                </span>
              </td>
              <td><span style="font-size: 12px; color: #475569;">{{ b.raw_filename }}</span></td>
              <td>{{ b.uploaded_at | date:'mediumDate' }} {{ b.uploaded_at | date:'shortTime' }}</td>
              <td style="text-align: right;">
                <div style="display: inline-flex; gap: 6px; align-items: center;">
                  <a
                    class="btn secondary"
                    [routerLink]="['/cost/sewtech/report', b.id]"
                    style="padding: 4px 10px; font-size: 12px;"
                  >
                    Review
                  </a>
                  <a
                    *ngIf="b.status === 'finalized'"
                    class="btn"
                    [routerLink]="['/cost/sewtech/final', b.id]"
                    style="padding: 4px 10px; font-size: 12px; background: #059669; border-color: #047857;"
                  >
                    Summary
                  </a>
                  <button
                    class="btn secondary"
                    (click)="api.downloadMerged(b.id)"
                    style="padding: 4px 10px; font-size: 12px;"
                    title="Download Merged Excel"
                  >
                    ⬇ Merged
                  </button>
                  <button
                    class="btn-delete"
                    (click)="deleteBatch(b)"
                    title="Delete this Sewtech batch and history"
                  >
                    Delete
                  </button>
                </div>
              </td>
            </tr>
            <tr *ngIf="batches.length === 0">
              <td colspan="7" style="text-align: center; color: #94a3b8; padding: 32px;">
                No Sewtech batches uploaded yet. Click "+ New Weekly Upload" to start.
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `,
})
export class SewtechBatchesComponent implements OnInit {
  batches: any[] = [];

  constructor(public api: SewtechApiService) {}

  ngOnInit() {
    this.loadBatches();
  }

  loadBatches() {
    this.api.listBatches().subscribe((res) => {
      this.batches = res || [];
    });
  }

  deleteBatch(b: any) {
    const msg = `Are you sure you want to delete Sewtech Batch #${b.id} (${b.week_label})?\n\nThis will permanently delete the batch and all calculated rows and summaries.`;
    if (!confirm(msg)) return;

    this.api.deleteBatch(b.id).subscribe({
      next: () => {
        this.loadBatches();
      },
      error: (err) => {
        alert(err?.error?.error || 'Failed to delete Sewtech batch.');
      },
    });
  }
}
