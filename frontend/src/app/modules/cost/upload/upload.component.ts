import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { ApiService } from '../core/services/api.service';

@Component({
  selector: 'app-upload',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="container">
      <div class="card">
        <div style="display: flex; align-items: flex-start; gap: 14px; margin-bottom: 20px;">
          <div style="width: 42px; height: 42px; border-radius: 10px; background: linear-gradient(135deg, #e6f6f7 0%, #f7ebfa 100%); display: flex; align-items: center; justify-content: center; font-size: 20px; flex-shrink: 0;">
            📤
          </div>
          <div>
            <h2 style="margin: 0 0 4px 0;">India Factory — Upload Raw Weekly Report</h2>
            <p style="margin: 0; color: #64748b; font-size: 13.5px;">
              Upload the raw export downloaded from the ERP. The automated pipeline will forward-fill missing styles, resolve buyer groups, drop low-quantity samples (≤ 1), fix zero SAM/Efficiency records, and compute machine requirements.
            </p>
          </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; margin: 24px 0 16px 0;">
          <div>
            <label style="display: block; margin-bottom: 6px;">Week Label <span style="color:#ef4444;">*</span></label>
            <input
              type="text"
              [(ngModel)]="weekLabel"
              placeholder="e.g. 21 Sep to 27 Sep"
              style="width: 100%;"
              required
            />
          </div>

          <div>
            <label style="display: block; margin-bottom: 6px;">Working Days This Week <span style="color:#ef4444;">*</span></label>
            <input
              type="number"
              [(ngModel)]="workingDays"
              min="1"
              max="7"
              style="width: 100%;"
              required
            />
          </div>

          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 6px;">Raw ERP Excel File (.xls / .xlsx) <span style="color:#ef4444;">*</span></label>
            <div style="border: 2px dashed #cbd5e1; border-radius: 10px; padding: 24px; text-align: center; background: #fafbfc; transition: all 0.2s ease;">
              <input
                type="file"
                (change)="onFile($event)"
                accept=".xls,.xlsx"
                id="rawFileInput"
                style="display: none;"
                required
              />
              <label for="rawFileInput" class="btn secondary" style="cursor: pointer; margin-right: 12px;">
                Choose Excel File
              </label>
              <span *ngIf="file" style="font-weight: 600; color: #0f172a;">
                📄 {{ file.name }} <span style="color: #64748b; font-weight: normal;">({{ (file.size / 1024).toFixed(1) }} KB)</span>
              </span>
              <span *ngIf="!file" style="color: #64748b; font-size: 13px;">
                or click to select your raw Excel workbook
              </span>
            </div>
          </div>
        </div>

        <div style="margin-top: 24px; display: flex; align-items: center; gap: 16px;">
          <button
            class="btn"
            [disabled]="!file || !weekLabel.trim() || !workingDays || workingDays < 1 || uploading"
            (click)="upload()"
          >
            {{ uploading ? 'Processing Data Pipeline…' : 'Upload & Process India Data' }}
          </button>
        </div>

        <div *ngIf="error" style="margin-top: 16px; padding: 12px 16px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; color: #b91c1c; font-size: 13.5px;">
          <strong>⚠ Error:</strong> {{ error }}
        </div>
      </div>
    </div>
  `,
})
export class UploadComponent {
  file: File | null = null;
  weekLabel = '';
  workingDays: number | null = 5;
  uploading = false;
  error = '';

  constructor(private api: ApiService, private router: Router) {}

  onFile(evt: Event) {
    const input = evt.target as HTMLInputElement;
    this.file = input.files?.[0] || null;
  }

  upload() {
    if (!this.file || !this.weekLabel.trim() || !this.workingDays || this.workingDays < 1) {
      this.error = 'Please fill in all mandatory fields: Week label, Working days (1-7), and choose a raw file.';
      return;
    }
    this.uploading = true;
    this.error = '';
    this.api.uploadRaw(this.file, this.weekLabel.trim(), Number(this.workingDays)).subscribe({
      next: (res) => {
        this.uploading = false;
        this.router.navigate(['/cost/india/batches', res.batch_id]);
      },
      error: (err) => {
        this.uploading = false;
        this.error = err?.error?.error || 'Upload failed. Check the backend is running and the file format matches the expected export.';
      },
    });
  }
}
