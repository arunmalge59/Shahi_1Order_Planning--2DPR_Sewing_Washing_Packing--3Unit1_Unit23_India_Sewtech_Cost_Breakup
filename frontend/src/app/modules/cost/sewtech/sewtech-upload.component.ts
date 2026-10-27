import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { SewtechApiService } from '../core/services/sewtech-api.service';

@Component({
  selector: 'app-sewtech-upload',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="container">
      <div class="card">
        <div style="display: flex; align-items: flex-start; gap: 14px; margin-bottom: 20px;">
          <div style="width: 42px; height: 42px; border-radius: 10px; background: linear-gradient(135deg, #f7ebfa 0%, #e6f6f7 100%); display: flex; align-items: center; justify-content: center; font-size: 20px; flex-shrink: 0;">
            📥
          </div>
          <div>
            <h2 style="margin: 0 0 4px 0;">Sewtech Factory — Upload Raw Daily SAM Production Report</h2>
            <p style="margin: 0; color: #64748b; font-size: 13.5px;">
              Upload the multi-tab daily production workbook (SFL#10 and SFL#11). The system automatically scans daily tabs, eliminates trainee rows, identifies missing dates, and prepares merged line metrics.
            </p>
          </div>
        </div>

        <!-- Step 1: Upload File -->
        <div style="margin: 20px 0; padding: 28px; border: 2px dashed #cbd5e1; border-radius: 10px; text-align: center; background: #fafbfc; transition: all 0.2s ease;">
          <input
            type="file"
            (change)="onFileSelected($event)"
            accept=".xlsx,.xlsm,.xls"
            id="sewtechFileInput"
            style="display: none;"
          />
          <label for="sewtechFileInput" class="btn secondary" style="cursor: pointer; margin-right: 12px;">
            Choose Sewtech File (.xlsx, .xlsm, .xls)
          </label>
          <span *ngIf="file" style="font-weight: 600; color: #0f172a;">
            📄 {{ file.name }} <span style="color: #64748b; font-weight: normal;">({{ (file.size / 1024).toFixed(1) }} KB)</span>
          </span>
          <span *ngIf="!file" style="color: #64748b; font-size: 13px;">
            or click to select your Sewtech workbook
          </span>

          <div style="margin-top: 18px;" *ngIf="file && !uploadResult">
            <button class="btn" [disabled]="uploading" (click)="uploadRawFile()">
              {{ uploading ? 'Scanning Daily Tabs…' : 'Scan Workbook Tabs' }}
            </button>
          </div>
        </div>

        <div *ngIf="error" style="margin-top: 16px; padding: 12px 16px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; color: #b91c1c; font-size: 13.5px;">
          <strong>⚠ Error:</strong> {{ error }}
        </div>

        <!-- Step 2: Date Range Selection & Detected Tabs -->
        <div *ngIf="uploadResult" style="margin-top: 28px; border-top: 1px solid #e2e8f0; padding-top: 24px;">
          <h3 style="margin-bottom: 12px;">2. Select Date Range & Confirm Working Days</h3>

          <!-- Detected tabs chips -->
          <div style="margin-bottom: 20px;">
            <label style="font-size: 12.5px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.04em;">
              Detected Daily Tabs in File:
            </label>
            <div style="display: flex; gap: 8px; flex-wrap: wrap; margin-top: 8px;">
              <span
                *ngFor="let tab of uploadResult.valid_tabs"
                class="badge green"
                style="padding: 4px 12px;"
                [title]="'Sheet: ' + tab.sheet_name + ' (' + tab.date + ')'"
              >
                ✓ {{ tab.date | date:'dd-MMM' }}
              </span>
              <span
                *ngFor="let ign of uploadResult.ignored_tabs"
                class="badge"
                style="background: #f1f5f9; color: #64748b; padding: 4px 10px;"
                [title]="'Ignored: ' + ign.reason"
              >
                {{ ign.sheetName }} (ignored)
              </span>
            </div>
          </div>

          <!-- Date pickers -->
          <div style="display: flex; gap: 20px; flex-wrap: wrap; align-items: flex-end; margin-bottom: 20px;">
            <div>
              <label style="display: block; margin-bottom: 4px;">Start Date <span style="color:#ef4444;">*</span></label>
              <input type="date" [(ngModel)]="startDate" (change)="onRangeChange()" />
            </div>

            <div>
              <label style="display: block; margin-bottom: 4px;">End Date <span style="color:#ef4444;">*</span></label>
              <input type="date" [(ngModel)]="endDate" (change)="onRangeChange()" />
            </div>

            <div>
              <label style="display: block; margin-bottom: 4px;">Working Days <span style="color:#ef4444;">*</span></label>
              <input type="number" [(ngModel)]="workingDays" min="1" max="31" style="width: 100px;" />
            </div>

            <button
              class="btn"
              [disabled]="processing || !startDate || !endDate || !workingDays || workingDays < 1"
              (click)="processRange()"
            >
              {{ processing ? 'Processing Range…' : 'Process Range & Build Merged Sheet' }}
            </button>
          </div>

          <!-- Range info display -->
          <div *ngIf="rangeInfo" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 18px;">
            <div style="display: flex; gap: 24px; flex-wrap: wrap; font-size: 13px;">
              <div>
                <strong style="color: #059669;">Found working tabs ({{ rangeInfo.found.length }}):</strong> {{ rangeInfo.found.join(', ') || 'None' }}
              </div>
              <div *ngIf="rangeInfo.missing.length > 0" style="color: #d97706;">
                <strong>Missing calendar dates in range ({{ rangeInfo.missing.length }}):</strong> {{ rangeInfo.missing.join(', ') }}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  `,
})
export class SewtechUploadComponent {
  file: File | null = null;
  uploading = false;
  processing = false;
  error = '';
  uploadResult: any = null;

  startDate = '';
  endDate = '';
  workingDays = 6;
  rangeInfo: { found: string[]; missing: string[] } | null = null;

  constructor(private api: SewtechApiService, private router: Router) {}

  onFileSelected(evt: Event) {
    const input = evt.target as HTMLInputElement;
    this.file = input.files?.[0] || null;
    this.uploadResult = null;
    this.error = '';
  }

  uploadRawFile() {
    if (!this.file) return;
    this.uploading = true;
    this.error = '';

    this.api.uploadRaw(this.file).subscribe({
      next: (res) => {
        this.uploading = false;
        this.uploadResult = res;
        this.startDate = res.min_date;
        this.endDate = res.max_date;
        this.onRangeChange();
      },
      error: (err) => {
        this.uploading = false;
        this.error = err?.error?.error || 'Failed to upload and scan workbook.';
      },
    });
  }

  onRangeChange() {
    if (!this.uploadResult || !this.startDate || !this.endDate) return;

    const validTabs: any[] = this.uploadResult.valid_tabs || [];
    const foundTabs = validTabs.filter((t) => t.date >= this.startDate && t.date <= this.endDate);

    const foundDates = foundTabs.map((t) => {
      const parts = t.date.split('-');
      return `${parseInt(parts[2], 10)}-${parts[1]}`;
    });

    const allDatesInRange: string[] = [];
    const cur = new Date(this.startDate);
    const end = new Date(this.endDate);
    while (cur <= end) {
      const d = String(cur.getDate()).padStart(2, '0');
      const m = String(cur.getMonth() + 1).padStart(2, '0');
      const y = cur.getFullYear();
      allDatesInRange.push(`${y}-${m}-${d}`);
      cur.setDate(cur.getDate() + 1);
    }

    const foundSet = new Set(foundTabs.map((t) => t.date));
    const missing = allDatesInRange
      .filter((d) => !foundSet.has(d))
      .map((d) => {
        const parts = d.split('-');
        return `${parseInt(parts[2], 10)}-${parts[1]}`;
      });

    this.rangeInfo = { found: foundDates, missing };
    this.workingDays = foundTabs.length;
  }

  processRange() {
    if (!this.uploadResult || !this.startDate || !this.endDate || !this.workingDays) return;
    this.processing = true;
    this.error = '';

    this.api.processBatch(this.uploadResult.batch_id, {
      start_date: this.startDate,
      end_date: this.endDate,
      working_days: Number(this.workingDays),
    }).subscribe({
      next: () => {
        this.processing = false;
        this.router.navigate(['/cost/sewtech/report', this.uploadResult.batch_id]);
      },
      error: (err) => {
        this.processing = false;
        this.error = err?.error?.error || 'Failed to process batch.';
      },
    });
  }
}
