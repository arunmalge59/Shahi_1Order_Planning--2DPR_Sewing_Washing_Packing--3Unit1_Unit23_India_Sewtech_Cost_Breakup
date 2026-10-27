import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { RouterLink } from '@angular/router';
import { ApiService } from '../core/services/api.service';

@Component({
  selector: 'app-batch-list',
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <div class="container">
      <div class="card">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
          <div>
            <h2 style="margin: 0 0 4px 0;">India Weekly Batches</h2>
            <p style="margin: 0; color: #64748b; font-size: 13px;">Manage, review, and finalize India factory weekly upload batches.</p>
          </div>
          <a routerLink="/cost/india/upload" class="btn">+ New Weekly Upload</a>
        </div>

        <table>
          <thead>
            <tr>
              <th>ID</th>
              <th>Week</th>
              <th>Working Days</th>
              <th>Status</th>
              <th>Uploaded</th>
              <th style="text-align: right;">Actions</th>
            </tr>
          </thead>
          <tbody>
            <tr *ngFor="let b of batches">
              <td><strong>#{{ b.id }}</strong></td>
              <td><strong>{{ b.week_label }}</strong></td>
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
              <td>{{ b.uploaded_at | date:'mediumDate' }} {{ b.uploaded_at | date:'shortTime' }}</td>
              <td style="text-align: right;">
                <div style="display: inline-flex; gap: 8px; align-items: center;">
                  <a class="btn secondary" [routerLink]="['/cost/india/batches', b.id]" style="padding: 4px 12px; font-size: 12px;">
                    Open
                  </a>
                  <button
                    class="btn-delete"
                    (click)="deleteBatch(b)"
                    title="Delete this batch and history"
                  >
                    Delete
                  </button>
                </div>
              </td>
            </tr>
            <tr *ngIf="batches.length === 0">
              <td colspan="6" style="text-align: center; color: #94a3b8; padding: 32px;">
                No India weekly batches found. Click "+ New Weekly Upload" to process raw data.
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `,
})
export class BatchListComponent implements OnInit {
  batches: any[] = [];

  constructor(private api: ApiService) {}

  ngOnInit() {
    this.loadBatches();
  }

  loadBatches() {
    this.api.listBatches().subscribe((b) => (this.batches = b || []));
  }

  deleteBatch(b: any) {
    const msg = `Are you sure you want to delete Batch #${b.id} (${b.week_label})?\n\nThis will permanently remove the batch and all associated rows and logs.`;
    if (!confirm(msg)) return;

    this.api.deleteBatch(b.id).subscribe({
      next: () => {
        this.loadBatches();
      },
      error: (err) => {
        alert(err?.error?.error || 'Failed to delete batch.');
      },
    });
  }
}
