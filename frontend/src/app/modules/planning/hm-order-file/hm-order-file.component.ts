import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';

const API_BASE_URL = typeof window !== 'undefined' && window.location.hostname === 'localhost' && window.location.port === '4200' ? 'http://localhost:4000/planning' : '/planning';

interface HmSummary {
  totalRows: number;
  projectionRows: number;
  confirmedRows: number;
  generatedOrders: number;
  skippedRows: number;
  validProjectionRows?: number;
  frProvided: boolean;
  frRowCount: number;
  frMatchedRows: number;
  frMissingRows: number;
}

interface HmValidationError {
  sourceRowNumber: number;
  status: string;
  style: string;
  qty: string;
  projExFty: string;
  projPcd: string;
  reason: string;
  type: 'data' | 'fr-missing';
}

interface HmResult {
  csv: string;
  summary: HmSummary;
  preview: Record<string, string | number>[];
  validationErrors: HmValidationError[];
  frMissingStyles: { season: string; style: string }[];
}

@Component({
  selector: 'app-hm-order-file',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './hm-order-file.component.html',
  styleUrl: './hm-order-file.component.css',
})
export class HmOrderFileComponent {
  file: File | null = null;
  frFile: File | null = null;
  weekCode = 'W33';
  loading = false;
  errorMessage = '';
  result: HmResult | null = null;
  showValidationErrors = false;

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.file = input.files?.[0] || null;
    this.result = null;
    this.showValidationErrors = false;
    this.errorMessage = '';
  }

  // The FR file is entirely optional — selecting nothing here just means the
  // order file falls back to the raw style/description behavior, as before.
  onFrFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.frFile = input.files?.[0] || null;
    this.result = null;
    this.showValidationErrors = false;
    this.errorMessage = '';
  }

  clearFrFile(): void {
    this.frFile = null;
    this.result = null;
    this.showValidationErrors = false;
  }

  async createOrderFile(): Promise<void> {
    this.errorMessage = '';
    if (!this.file) {
      this.errorMessage = 'Please choose an H&M planning workbook.';
      return;
    }
    if (!/^W\d{1,2}$/i.test(this.weekCode.trim())) {
      this.errorMessage = 'Week Code must be in the format W33.';
      return;
    }

    this.loading = true;
    this.result = null;
    this.showValidationErrors = false;
    const formData = new FormData();
    formData.append('file', this.file);
    formData.append('weekCode', this.weekCode.trim());
    if (this.frFile) {
      formData.append('frFile', this.frFile);
    }

    try {
      const response = await fetch(`${API_BASE_URL}/api/hm/order-file`, { method: 'POST', body: formData });
      const body = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(body.error || `Server returned status ${response.status}`);
      this.result = body as HmResult;
    } catch (error: any) {
      this.errorMessage = error.message || 'Could not create the H&M order file.';
    } finally {
      this.loading = false;
    }
  }

  downloadOrderFile(): void {
    if (!this.result) return;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([this.result.csv], { type: 'text/csv;charset=utf-8' }));
    link.download = `hm-order-file-${this.weekCode.trim().toUpperCase()}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  get previewColumns(): string[] {
    return ['O.CODE', 'O.PROD', 'O.UDITEM', 'O.CUST', 'O^DD:1', 'O^DQ:1', 'O.CONTRACT_QTY', 'O.UDSEASON', 'O.UDPCD', 'O.STATUS'];
  }
}