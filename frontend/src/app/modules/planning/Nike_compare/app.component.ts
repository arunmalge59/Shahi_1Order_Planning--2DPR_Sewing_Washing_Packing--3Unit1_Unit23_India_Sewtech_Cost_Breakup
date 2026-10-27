import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx'; // npm install xlsx (if not already installed)

export interface ScheduleComparisonRecord {
  style: string;
  description: string;
  julyPcd: string;
  augustPcd: string;
  pcdDiff: string;
  julyExFactory: string;
  augustExFactory: string;
  exFactoryDiff: string;
  julyQty: number | null;
  augustQty: number | null;
  qtyDiff: number;
  qtyDiffStr: string;
  status: 'Same' | 'Changed' | 'Added' | 'Removed';
  buyDate?: string;
  itemNumber?: string;
}

export interface StyleResult {
  style: string;
  description: string;
  julyQty: number;
  augustQty: number;
  qtyDelta: number;
  pcdChanged: boolean;
  exFactoryChanged: boolean;
  presenceChange: string;
  status: 'HIGH' | 'MEDIUM' | 'LOW';
  julySchedule: any[];
  augustSchedule: any[];
  schedules: ScheduleComparisonRecord[];
}

export interface CompareSummary {
  totalStyles: number;
  changedStyles: number;
  qtyChanged: number;
  pcdChanged: number;
  exFactoryChanged: number;
  highPriority: number;
  mediumPriority: number;
  lowPriority: number;
  totalSchedules: number;
  sameSchedules: number;
  changedSchedules: number;
  addedSchedules: number;
  removedSchedules: number;
}

export interface CompareResult {
  summary: CompareSummary;
  styles: StyleResult[];
  allSchedules: ScheduleComparisonRecord[];
  runId?: number;
  dimensionRows?: DimensionComparisonRow[];
}

export interface DimensionComparisonRow {
  style: string;
  season: string;
  month: string;
  julyQty: number;
  augustQty: number;
  variation: number;
}

const API_BASE_URL = typeof window !== 'undefined' && window.location.hostname === 'localhost' && window.location.port === '4200' ? 'http://localhost:4000/planning' : '/planning';

@Component({
  selector: 'app-compare-snapshot',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  julyFile: File | null = null;
  augustFile: File | null = null;

  loading = false;
  errorMessage = '';

  result: CompareResult | null = null;

  // Filters
  selectedStyleFilter = 'ALL';
  descriptionFilter = '';
  scheduleStatusFilter: 'ALL' | 'Changed' | 'Same' | 'Added' | 'Removed' = 'ALL';

  newComparisonStyle = 'ALL';
  newComparisonSeason = 'All';
  newComparisonMonth = 'All';

  selectedStyleModal: StyleResult | null = null;

  onJulyFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    this.julyFile = input.files && input.files.length ? input.files[0] : null;
  }

  onAugustFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    this.augustFile = input.files && input.files.length ? input.files[0] : null;
  }

  async compare() {
    this.errorMessage = '';

    if (!this.julyFile || !this.augustFile) {
      this.errorMessage = 'Please choose both the older snapshot and the newer snapshot CSV files.';
      return;
    }

    const formData = new FormData();
    formData.append('julyFile', this.julyFile);
    formData.append('augustFile', this.augustFile);

    this.loading = true;
    this.result = null;

    try {
      const response = await fetch(`${API_BASE_URL}/api/compare`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const body = await response.json().catch(() => ({}));
        throw new Error(body.error || `Server returned status ${response.status}`);
      }

      this.result = await response.json();
    } catch (err: any) {
      this.errorMessage = err.message || 'Something went wrong while comparing the files.';
    } finally {
      this.loading = false;
    }
  }

  get availableStyles(): string[] {
    if (!this.result || !this.result.styles) return [];
    const set = new Set<string>();
    for (const s of this.result.styles) {
      if (s.style) set.add(s.style);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }

  get newComparisonRows(): DimensionComparisonRow[] {
    if (!this.result) return [];
    return (this.result.dimensionRows || []).filter((row) =>
      (this.newComparisonStyle === 'ALL' || row.style === this.newComparisonStyle) &&
      (this.newComparisonSeason === 'All' || row.season === this.newComparisonSeason) &&
      (this.newComparisonMonth === 'All' || row.month === this.newComparisonMonth)
    ).sort((a, b) => a.style.localeCompare(b.style) || a.season.localeCompare(b.season) || a.month.localeCompare(b.month));
  }

  get newComparisonStyles(): string[] {
    return Array.from(new Set((this.result?.dimensionRows || []).map((row) => row.style))).sort();
  }

  get newComparisonSeasons(): string[] {
    return Array.from(new Set((this.result?.dimensionRows || []).map((row) => row.season))).sort();
  }

  get newComparisonMonths(): string[] {
    const order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'Unknown'];
    return Array.from(new Set((this.result?.dimensionRows || []).map((row) => row.month))).sort((a, b) => order.indexOf(a) - order.indexOf(b));
  }

  get newComparisonGrandTotal(): { julyQty: number; augustQty: number; variation: number } {
    return this.newComparisonRows.reduce((total, row) => ({
      julyQty: total.julyQty + row.julyQty,
      augustQty: total.augustQty + row.augustQty,
      variation: total.variation + row.variation,
    }), { julyQty: 0, augustQty: 0, variation: 0 });
  }

  get filteredSchedules(): ScheduleComparisonRecord[] {
    if (!this.result) return [];
    let list = this.result.allSchedules || [];

    // Style filter
    if (this.selectedStyleFilter !== 'ALL') {
      list = list.filter((r) => r.style === this.selectedStyleFilter);
    }

    // Description filter
    if (this.descriptionFilter.trim() !== '') {
      const term = this.descriptionFilter.trim().toLowerCase();
      list = list.filter((r) => r.description && r.description.toLowerCase().includes(term));
    }

    // Status filter
    if (this.scheduleStatusFilter === 'Changed') {
      // "Changed" is treated as a combined bucket: Changed + Added + Removed.
      // "Same" schedules are always excluded from this view.
      list = list.filter((r) => r.status !== 'Same');
    } else if (this.scheduleStatusFilter !== 'ALL') {
      list = list.filter((r) => r.status === this.scheduleStatusFilter);
    }

    return list;
  }

  /** Totals for whatever is currently showing in the Schedule-Level table (respects all active filters). */
  get filteredQtyTotals(): { julyQty: number; augustQty: number; qtyDiff: number } {
    return this.filteredSchedules.reduce(
      (totals, r) => ({
        julyQty: totals.julyQty + (r.julyQty || 0),
        augustQty: totals.augustQty + (r.augustQty || 0),
        qtyDiff: totals.qtyDiff + (r.qtyDiff || 0),
      }),
      { julyQty: 0, augustQty: 0, qtyDiff: 0 }
    );
  }

  /** Net qty change across every schedule record in the comparison (unfiltered) — shown as a top KPI. */
  get overallQtyChange(): number {
    if (!this.result) return 0;
    return (this.result.allSchedules || []).reduce((sum, r) => sum + (r.qtyDiff || 0), 0);
  }

  get filteredStyleSummaries(): StyleResult[] {
    if (!this.result) return [];
    let list = this.result.styles || [];

    if (this.selectedStyleFilter !== 'ALL') {
      list = list.filter((s) => s.style === this.selectedStyleFilter);
    }

    if (this.descriptionFilter.trim() !== '') {
      const term = this.descriptionFilter.trim().toLowerCase();
      list = list.filter((s) => s.description && s.description.toLowerCase().includes(term));
    }

    return list;
  }

  setScheduleStatusFilter(status: 'ALL' | 'Changed' | 'Same' | 'Added' | 'Removed') {
    this.scheduleStatusFilter = status;
  }

  filterByStyle(style: string) {
    this.selectedStyleFilter = style;
  }

  openDetail(styleObj: StyleResult) {
    this.selectedStyleModal = styleObj;
  }

  closeDetail() {
    this.selectedStyleModal = null;
  }

  resetFilters() {
    this.selectedStyleFilter = 'ALL';
    this.descriptionFilter = '';
    this.scheduleStatusFilter = 'ALL';
  }

  statusBadge(status: string): string {
    if (status === 'HIGH') return '🔴 HIGH';
    if (status === 'MEDIUM') return '🟠 MEDIUM';
    if (status === 'LOW') return '🟢 LOW';
    return status;
  }

  /**
   * Downloads the currently filtered Schedule-Level rows as an Excel (.xlsx) file.
   * Whatever is on screen (Style / Description / Status filters applied) is exactly
   * what gets exported — e.g. filter to "Changed" to get a Sep-month changed-only report.
   */
  downloadFilteredReport() {
    const rows = this.filteredSchedules;
    if (!rows.length) return;

    const exportRows = rows.map((r) => ({
      Style: r.style,
      Description: r.description,
      'July EX Factory': r.julyExFactory,
      'August EX Factory': r.augustExFactory,
      'EX Factory Diff': r.exFactoryDiff,
      'July Qty': r.julyQty ?? '',
      'August Qty': r.augustQty ?? '',
      'Qty Diff': r.qtyDiff,
      Status: r.status,
    }));

    // Append an overall total row at the bottom of the export
    const totals = this.filteredQtyTotals;
    exportRows.push({
      Style: '',
      Description: 'TOTAL',
      'July EX Factory': '',
      'August EX Factory': '',
      'EX Factory Diff': '',
      'July Qty': totals.julyQty,
      'August Qty': totals.augustQty,
      'Qty Diff': totals.qtyDiff,
      Status: '',
    } as any);

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    worksheet['!cols'] = [
      { wch: 10 }, { wch: 32 }, { wch: 16 }, { wch: 16 }, { wch: 14 },
      { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    ];

    const workbook = XLSX.utils.book_new();
    const sheetLabel = this.scheduleStatusFilter === 'ALL' ? 'All' : this.scheduleStatusFilter;
    XLSX.utils.book_append_sheet(workbook, worksheet, `${sheetLabel} Schedules`.substring(0, 31));

    const today = new Date();
    const stamp = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(workbook, `Nike_Schedule_${sheetLabel}_Report_${stamp}.xlsx`);
  }
}