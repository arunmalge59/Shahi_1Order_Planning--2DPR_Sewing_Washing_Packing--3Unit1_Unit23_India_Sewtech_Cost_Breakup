import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

export interface TableCellMonth {
  prev: number;
  curr: number;
  diff: number;
}

export interface ComparisonRow {
  category: string;
  months: { [key: string]: TableCellMonth };
}

export interface ComparisonTotals {
  category: string;
  months: { [key: string]: TableCellMonth };
}

export interface ComparisonTable {
  id: string;
  title: string;
  rows: ComparisonRow[];
  totals: ComparisonTotals;
}

export interface MonthHeader {
  key: string;
  label: string;
}

export interface MonthOption {
  key: string;
  year: number;
  month: number;
  label: string;
}

export interface GrandTotals {
  prev: { total: number; confirmed: number; projections: number };
  curr: { total: number; confirmed: number; projections: number };
}

export interface SanityChecks {
  passed: boolean;
  errors: string[];
}

export interface ComparisonPreviewResult {
  monthHeaders: MonthHeader[];
  activeMonths: { key: string; year: number; month: number }[];
  categories: string[];
  tables: ComparisonTable[];
  prevLabel: string;
  currLabel: string;
  fromMonthKey: string;
  toMonthKey: string;
  fromLabel: string;
  toLabel: string;
  grandTotals: GrandTotals;
  sanityChecks: SanityChecks;
  warnings: string[];
  isSameFile?: boolean;
}

const API_BASE_URL =
  typeof window !== 'undefined' &&
  window.location.hostname === 'localhost' &&
  window.location.port === '4200'
    ? 'http://localhost:4000/planning'
    : '/planning';

@Component({
  selector: 'app-nike-category-comparison',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './nike-category-comparison.component.html',
  styleUrl: './nike-category-comparison.component.css',
})
export class NikeCategoryComparisonComponent {
  previousFile: File | null = null;
  currentFile: File | null = null;

  prevLabel: string = '';
  currLabel: string = '';

  monthsLoading = false;
  comparing = false;
  downloading = false;

  errorMessage = '';
  sameFileNotice = false;

  monthOptions: MonthOption[] = [];
  fromMonthKey: string = '';
  toMonthKey: string = '';

  previewResult: ComparisonPreviewResult | null = null;

  private formatFileDate(file: File): string {
    const d = new Date(file.lastModified || Date.now());
    const day = String(d.getDate()).padStart(2, '0');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = months[d.getMonth()];
    return `${day}-${month}`;
  }

  onPreviousFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.previousFile = input.files[0];
      this.prevLabel = this.formatFileDate(this.previousFile);
      this.onFilesChanged();
    }
  }

  onCurrentFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.currentFile = input.files[0];
      this.currLabel = this.formatFileDate(this.currentFile);
      this.onFilesChanged();
    }
  }

  private onFilesChanged(): void {
    this.previewResult = null;
    this.errorMessage = '';
    this.sameFileNotice = false;
    this.monthOptions = [];
    this.fromMonthKey = '';
    this.toMonthKey = '';

    if (this.previousFile && this.currentFile) {
      this.loadMonthOptions();
    }
  }

  async loadMonthOptions(): Promise<void> {
    if (!this.previousFile || !this.currentFile) return;

    this.monthsLoading = true;
    this.errorMessage = '';

    const formData = new FormData();
    formData.append('previousFile', this.previousFile);
    formData.append('currentFile', this.currentFile);

    try {
      const response = await fetch(`${API_BASE_URL}/api/nike-category-compare/months`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || `Server responded with ${response.status}`);
      }

      const data = await response.json();
      this.monthOptions = data.months || [];
      this.fromMonthKey = data.defaultFrom || (this.monthOptions[0]?.key ?? '');
      this.toMonthKey = data.defaultTo || (this.monthOptions[this.monthOptions.length - 1]?.key ?? '');
      this.sameFileNotice = !!data.isSameFile;
    } catch (err: any) {
      this.errorMessage = err.message || 'Failed to extract available months from the uploaded files.';
    } finally {
      this.monthsLoading = false;
    }
  }

  get isRangeInvalid(): boolean {
    if (!this.fromMonthKey || !this.toMonthKey) return false;
    return this.fromMonthKey > this.toMonthKey;
  }

  async compareSnapshots(): Promise<void> {
    if (!this.previousFile || !this.currentFile) return;
    if (this.isRangeInvalid) {
      this.errorMessage = 'From Month cannot be later than To Month.';
      return;
    }

    this.comparing = true;
    this.errorMessage = '';

    const formData = new FormData();
    formData.append('previousFile', this.previousFile);
    formData.append('currentFile', this.currentFile);
    formData.append('prevLabel', this.prevLabel.trim() || 'Previous');
    formData.append('currLabel', this.currLabel.trim() || 'Current');
    formData.append('fromMonth', this.fromMonthKey);
    formData.append('toMonth', this.toMonthKey);

    try {
      const response = await fetch(`${API_BASE_URL}/api/nike-category-compare/preview`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err.error || `Comparison failed with status ${response.status}`);
      }

      const data: ComparisonPreviewResult = await response.json();
      this.previewResult = data;
      this.sameFileNotice = !!data.isSameFile;
    } catch (err: any) {
      this.errorMessage = err.message || 'Failed to compare snapshots.';
    } finally {
      this.comparing = false;
    }
  }

  async downloadExcel(): Promise<void> {
    if (!this.previousFile || !this.currentFile) return;

    this.downloading = true;
    this.errorMessage = '';

    const formData = new FormData();
    formData.append('previousFile', this.previousFile);
    formData.append('currentFile', this.currentFile);
    formData.append('prevLabel', this.prevLabel.trim() || 'Previous');
    formData.append('currLabel', this.currLabel.trim() || 'Current');
    formData.append('fromMonth', this.fromMonthKey);
    formData.append('toMonth', this.toMonthKey);

    try {
      const response = await fetch(`${API_BASE_URL}/api/nike-category-compare/export`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err.error || 'Failed to download Excel.');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const safePrev = (this.prevLabel.trim() || 'Previous').replace(/[^a-zA-Z0-9_-]/g, '_');
      const safeCurr = (this.currLabel.trim() || 'Current').replace(/[^a-zA-Z0-9_-]/g, '_');
      a.download = `Nike_Category_Comparison_${safePrev}_vs_${safeCurr}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      this.errorMessage = err.message || 'Error exporting Excel file.';
    } finally {
      this.downloading = false;
    }
  }

  formatQty(val: number | null | undefined): string {
    if (val === 0 || val === null || val === undefined) return '-';
    return val.toLocaleString('en-IN');
  }

  formatDiff(val: number | null | undefined): string {
    if (val === 0 || val === null || val === undefined) return '-';
    return val.toLocaleString('en-IN');
  }
}
