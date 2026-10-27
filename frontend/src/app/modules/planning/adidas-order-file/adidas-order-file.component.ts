import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';

const API_BASE_URL = typeof window !== 'undefined' && window.location.hostname === 'localhost' && window.location.port === '4200' ? 'http://localhost:4000/planning' : '/planning';

interface AdidasSummary {
  totalRows: number;
  validRows: number;
  generatedOrders: number;
  skippedRows: number;
  totalContractQty: number;
  totalOrderQty: number;
}

interface AdidasValidationError {
  sourceRowNumber: number;
  style: string;
  qty: string;
  reason: string;
}

interface AdidasResult {
  csv: string;
  sheetName: string;
  summary: AdidasSummary;
  preview: Record<string, string | number>[];
  validationErrors: AdidasValidationError[];
}

@Component({
  selector: 'app-adidas-order-file',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './adidas-order-file.component.html',
  styleUrl: './adidas-order-file.component.css',
})
export class AdidasOrderFileComponent {
  file: File | null = null;
  weekCode = 'WK-33 Rev1';
  loading = false;
  errorMessage = '';
  result: AdidasResult | null = null;
  showValidationErrors = false;

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.file = input.files?.[0] || null;
    this.result = null;
    this.showValidationErrors = false;
    this.errorMessage = '';
  }

  async createOrderFile(): Promise<void> {
    this.errorMessage = '';
    if (!this.file) {
      this.errorMessage = 'Please choose an Adidas Order Chart workbook.';
      return;
    }

    this.loading = true;
    this.result = null;
    this.showValidationErrors = false;
    const formData = new FormData();
    formData.append('file', this.file);
    formData.append('weekCode', this.weekCode.trim() || 'WK-33 Rev1');

    try {
      const response = await fetch(`${API_BASE_URL}/api/adidas/order-file`, {
        method: 'POST',
        body: formData,
      });
      const body = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(body.error || `Server returned status ${response.status}`);
      this.result = body as AdidasResult;
    } catch (error: any) {
      this.errorMessage = error.message || 'Could not create the Adidas order file.';
    } finally {
      this.loading = false;
    }
  }

  downloadOrderFile(): void {
    if (!this.result) return;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([this.result.csv], { type: 'text/csv;charset=utf-8' }));
    const cleanWeek = this.weekCode.trim().replace(/[^a-z0-9_-]/gi, '_');
    link.download = `Orders-ADIDAS-${cleanWeek || 'generated'}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  get previewColumns(): string[] {
    return [
      'O.CODE',
      'O.PROD',
      'O.CUST',
      'O^DD:1',
      'O^DQ:1',
      'O.CONTRACT_QTY',
      'O.UDSEASON',
      'O.UDPCD',
      'O.STATUS',
      'O.Descrip',
      'O.Udconversion',
      'END',
    ];
  }
}
