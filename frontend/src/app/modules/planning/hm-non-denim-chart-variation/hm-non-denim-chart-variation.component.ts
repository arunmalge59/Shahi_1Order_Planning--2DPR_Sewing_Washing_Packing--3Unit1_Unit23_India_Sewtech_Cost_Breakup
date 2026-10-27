import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';
import { HmNonDenimChartVariationService } from './hm-non-denim-chart-variation.service';

@Component({
  selector: 'app-hm-non-denim-chart-variation',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './hm-non-denim-chart-variation.component.html',
  styleUrl: './hm-non-denim-chart-variation.component.css',
})
export class HmNonDenimChartVariationComponent {
  previousWeekFile: File | null = null;
  currentWeekFile: File | null = null;
  monthOptions: string[] = [];
  startMonth = '';
  endMonth = '';
  loading = false;
  errorMessage = '';
  successMessage = '';

  constructor(
    private chartVariationService: HmNonDenimChartVariationService,
    private router: Router
  ) {}

  async onPreviousWeekFileSelected(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    this.previousWeekFile = input.files?.[0] || null;
    this.errorMessage = '';
    this.successMessage = '';
    await this.refreshMonthOptions();
  }

  async onCurrentWeekFileSelected(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    this.currentWeekFile = input.files?.[0] || null;
    this.errorMessage = '';
    this.successMessage = '';
    await this.refreshMonthOptions();
  }

  private async refreshMonthOptions(): Promise<void> {
    if (!this.previousWeekFile || !this.currentWeekFile) {
      this.monthOptions = [];
      this.startMonth = '';
      this.endMonth = '';
      return;
    }

    try {
      const months = await this.chartVariationService.getAvailableMonths(this.previousWeekFile, this.currentWeekFile);
      this.monthOptions = months;
      this.startMonth = months[0] || '';
      this.endMonth = months[months.length - 1] || '';
    } catch (error: any) {
      this.monthOptions = [];
      this.startMonth = '';
      this.endMonth = '';
      this.errorMessage = error.message || 'Unable to load available months.';
    }
  }

  isRangeValid(): boolean {
    if (!this.startMonth || !this.endMonth) {
      return true;
    }
    return this.monthOptions.indexOf(this.startMonth) <= this.monthOptions.indexOf(this.endMonth);
  }

  async compareCharts(): Promise<void> {
    this.errorMessage = '';
    this.successMessage = '';

    if (!this.previousWeekFile || !this.currentWeekFile) {
      this.errorMessage = 'Please select both previous week and current week Excel files.';
      return;
    }

    if (!this.isRangeValid()) {
      this.errorMessage = 'Start Month cannot be later than End Month.';
      return;
    }

    this.loading = true;

    try {
      const blob = await this.chartVariationService.compareWeeklyCharts(
        this.previousWeekFile,
        this.currentWeekFile,
        this.startMonth,
        this.endMonth
      );

      this.successMessage = 'Chart variation created successfully.';

      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = this.generateFilename();
      link.click();
      URL.revokeObjectURL(link.href);
    } catch (error: any) {
      this.errorMessage = error.message || 'Failed to generate chart variation file.';
    } finally {
      this.loading = false;
    }
  }

  private generateFilename(): string {
    const previousWeekCode = this.extractWeekCode(this.previousWeekFile?.name || '');
    const currentWeekCode = this.extractWeekCode(this.currentWeekFile?.name || '');
    const base = previousWeekCode && currentWeekCode
      ? `H&M_Chart_Variation_Non-Denim_${previousWeekCode.toUpperCase()}_vs_${currentWeekCode.toUpperCase()}`
      : 'H&M_Chart_Variation_Non-Denim';

    if (!this.startMonth || !this.endMonth || this.startMonth === this.monthOptions[0] && this.endMonth === this.monthOptions[this.monthOptions.length - 1]) {
      return `${base}.xlsx`;
    }

    return `${base}_(${this.startMonth}-${this.endMonth}).xlsx`;
  }

  private extractWeekCode(filename: string): string | null {
    const match = filename.match(/WK-?(\d{6})/i);
    if (match) {
      const digits = match[1];
      const weekNum = digits.slice(-2);
      return `wk-${weekNum}`;
    }
    return null;
  }

  goBack(): void {
    this.router.navigate(['/hm-order-file-non-denim']);
  }
}
