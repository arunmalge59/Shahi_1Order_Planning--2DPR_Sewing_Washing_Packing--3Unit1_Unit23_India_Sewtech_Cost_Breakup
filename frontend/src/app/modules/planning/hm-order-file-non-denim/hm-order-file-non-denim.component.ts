import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';

const API_BASE_URL = typeof window !== 'undefined' && window.location.hostname === 'localhost' && window.location.port === '4200' ? 'http://localhost:4000/planning' : '/planning';

interface HmNonDenimSummary {
  totalCpRows: number;
  totalMarketingRows: number;
  generatedOrders: number;
  unmatchedCpRows: number;
  unmatchedMarketingRows: number;
  validationErrors: number;
}

interface HmNonDenimValidationError {
  sourceRowNumber: number;
  reason: string;
  style?: string;
  season?: string;
  itemNo?: string;
}

interface HmNonDenimResult {
  csv: string;
  summary: HmNonDenimSummary;
  preview: Record<string, string | number>[];
  validationErrors: HmNonDenimValidationError[];
}

@Component({
  selector: 'app-hm-order-file-non-denim',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './hm-order-file-non-denim.component.html',
  styleUrl: './hm-order-file-non-denim.component.css',
})
export class HmOrderFileNonDenimComponent {
  marketingFile: File | null = null;
  cpFile: File | null = null;
  previousCpFile: File | null = null;
  weekCode = 'wk-34';
  weekCodePattern = /^wk-\d{1,2}$/i;
  loading = false;
  cpTemplateLoading = false;
  cpTemplateGenerated = false;
  cpTemplateMessage = '';
  errorMessage = '';
  result: HmNonDenimResult | null = null;
  showValidationErrors = false;

  constructor(private router: Router) {}

  onMarketingFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.marketingFile = input.files?.[0] || null;
    this.result = null;
    this.cpTemplateGenerated = false;
    this.cpTemplateMessage = '';
    this.showValidationErrors = false;
    this.errorMessage = '';
  }

  onCpFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.cpFile = input.files?.[0] || null;
    this.result = null;
    this.showValidationErrors = false;
    this.errorMessage = '';
  }

  onPreviousCpFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.previousCpFile = input.files?.[0] || null;
    this.cpTemplateGenerated = false;
    this.cpTemplateMessage = '';
    this.errorMessage = '';
  }

  async downloadHmCpFile(): Promise<void> {
    this.errorMessage = '';
    this.cpTemplateMessage = '';
    this.cpTemplateGenerated = false;
    if (!this.marketingFile) {
      this.errorMessage = 'Please choose an Initial Order Chart (Marketing) file.';
      return;
    }

    this.cpTemplateLoading = true;
    const formData = new FormData();
    formData.append('marketingFile', this.marketingFile);
    if (this.previousCpFile) formData.append('previousCpFile', this.previousCpFile);

    try {
      const response = await fetch(`${API_BASE_URL}/api/hm/order-file-non-denim/cp-template`, {
        method: 'POST',
        body: formData,
      });
      if (!response.ok) {
        const body = await response.json().catch(() => ({}));
        throw new Error(body.error || `Server returned status ${response.status}`);
      }
      const generatedFile = await response.blob();
      const link = document.createElement('a');
      link.href = URL.createObjectURL(generatedFile);
      link.download = 'hm-non-denim-cp-file.xlsx';
      link.click();
      URL.revokeObjectURL(link.href);
      
      // Check for summary counts in headers
      const autoFilledCount = response.headers.get('X-Auto-Filled');
      const itemKnownCount = response.headers.get('X-Flagged-Item-Known');
      const newItemCount = response.headers.get('X-Flagged-New-Item');
      const flaggedCount = response.headers.get('X-Flagged-For-Review');

      this.cpTemplateGenerated = true;
      if (flaggedCount && parseInt(flaggedCount) > 0) {
        this.cpTemplateMessage = `✓ H&M CP file generated: ${autoFilledCount || 0} rows auto-filled. ⚠️ ${flaggedCount} rows flagged for manual review (${itemKnownCount || 0} orange - item known, ${newItemCount || 0} yellow - brand new item). Please review highlighted Unit cells.`;
      } else {
        this.cpTemplateMessage = `✓ H&M CP file generated and downloaded (${autoFilledCount || 0} rows auto-filled).`;
      }
    } catch (error: any) {
      this.errorMessage = error.message || 'Could not create the H&M CP file.';
    } finally {
      this.cpTemplateLoading = false;
    }
  }

  async createOrderFile(): Promise<void> {
    this.errorMessage = '';
    if (!this.marketingFile) {
      this.errorMessage = 'Please choose an Initial Order Chart (Marketing) file.';
      return;
    }
    if (!this.cpFile) {
      this.errorMessage = 'Please choose an H&M CP File.';
      return;
    }
    if (!this.weekCodePattern.test(this.weekCode.trim())) {
      this.errorMessage = 'Week Code must be in the format wk-34.';
      return;
    }

    this.loading = true;
    this.result = null;
    this.showValidationErrors = false;
    const formData = new FormData();
    formData.append('marketingFile', this.marketingFile);
    formData.append('cpFile', this.cpFile);
    formData.append('weekCode', this.weekCode.trim());
    if (this.previousCpFile) {
      formData.append('previousCpFile', this.previousCpFile);
    }

    try {
      const response = await fetch(`${API_BASE_URL}/api/hm/order-file-non-denim`, {
        method: 'POST',
        body: formData,
      });
      const body = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(body.error || `Server returned status ${response.status}`);
      this.result = body as HmNonDenimResult;
    } catch (error: any) {
      this.errorMessage = error.message || 'Could not create the H&M Non-Denim order file.';
    } finally {
      this.loading = false;
    }
  }

  downloadOrderFile(): void {
    if (!this.result) return;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([this.result.csv], { type: 'text/csv;charset=utf-8' }));
    link.download = `hm-order-file-non-denim-${this.weekCode.trim().toUpperCase()}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  get previewColumns(): string[] {
    return ['O.CODE', 'O.PROD', 'O.CUST', 'O^DD:1', 'O^DQ:1', 'O.CONTRACT_QTY', 'O.UDSEASON', 'O.UDPCD', 'O.STATUS', 'O.Descrip'];
  }

  navigateToChartVariation(): void {
    this.router.navigate(['/hm-non-denim-chart-variation']);
  }
}


