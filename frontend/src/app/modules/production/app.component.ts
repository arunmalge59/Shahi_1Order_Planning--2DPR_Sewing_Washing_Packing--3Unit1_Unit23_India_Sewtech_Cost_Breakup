import { Component, EventEmitter, Input, OnDestroy, OnInit, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpErrorResponse } from '@angular/common/http';
import { switchMap } from 'rxjs';
import {
  LinePlanService,
  ProcessSummary,
  FillActualsSummary,
  DprReportUpload,
  OrderSummarySummary,
  JobRecord,
} from './services/lineplan.service';

type ViewState = 'upload' | 'loading' | 'result';
type ActualsState = 'idle' | 'loading' | 'done';
type OrderSummaryState = 'idle' | 'loading' | 'done';
type CalendarCell = {
  day: number | null;
  iso: string | null;
  isWeekend: boolean;
  isToday: boolean;
  isAvailable: boolean;
};
type Unit23DprReportEntry = {
  file: File | null;
  reportDate: string;
};
type DirectUploadTarget =
  | 'unit1-workbook'
  | 'unit1-sewing'
  | 'unit1-wash-pack'
  | 'unit23-workbook'
  | 'unit23-sewing'
  | 'unit23-dpr';

@Component({
  selector: 'app-production-workspace',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css', './production-workspace.css'],
})
export class AppComponent implements OnInit, OnDestroy {
  @Output() workflowChange = new EventEmitter<{ unit: 'unit1' | 'unit23'; mode: 'full' | 'daily' }>();
  @Input() externalNavigation = false;
  @Input() set workflowUnit(unit: 'unit1' | 'unit23') {
    this.activeUnit = unit;
  }
  @Input() set workflowMode(mode: 'full' | 'daily') {
    this.pageMode = mode;
  }

  view: ViewState = 'upload';
  dragOver = false;

  selectedFile: File | null = null;
  errorMessage: string | null = null;

  summary: ProcessSummary = {};
  downloadUrl = '';
  planId = '';
  downloadError: string | null = null;
  downloading = false;

  availableDates: string[] = [];
  selectedYear = '';
  selectedMonth = '';
  excludedDates = new Set<string>();

  // Step 1.5: Order Summary Upload & Auto-Fill Order Summary
  orderSummaryState: OrderSummaryState = 'idle';
  orderSummaryDragOver = false;
  orderSummaryFile: File | null = null;
  orderSummaryError: string | null = null;
  orderSummarySummary: OrderSummarySummary = {};
  orderSummaryDownloadUrl = '';

  activeUnit: 'unit1' | 'unit23' = 'unit1';
  pageMode: 'full' | 'daily' = 'full';
  mobileNavOpen = false;

  setNav(unit: 'unit1' | 'unit23', page: 'full' | 'daily'): void {
    this.activeUnit = unit;
    this.pageMode = page;
    this.mobileNavOpen = false;
    this.workflowChange.emit({ unit, mode: page });
  }

  toggleMobileNav(): void {
    this.mobileNavOpen = !this.mobileNavOpen;
  }

  fileExtension(file: File): string {
    const extensionStart = file.name.lastIndexOf('.');
    return extensionStart >= 0 ? file.name.slice(extensionStart + 1).toUpperCase() : 'FILE';
  }

  // ==================== UNIT 23 STATE & METHODS ====================
  unit23View: ViewState = 'upload';
  unit23DragOver = false;
  unit23SelectedFile: File | null = null;
  unit23ErrorMessage: string | null = null;
  unit23Summary: ProcessSummary = {};
  unit23DownloadUrl = '';
  unit23PlanId = '';
  unit23DownloadError: string | null = null;
  unit23Downloading = false;
  unit23AvailableDates: string[] = [];
  unit23SelectedYear = '';
  unit23SelectedMonth = '';
  unit23ExcludedDates = new Set<string>();

  // Unit 23 Step 2: Order Summary Upload
  unit23OrderSummaryState: OrderSummaryState = 'idle';
  unit23OrderSummaryDragOver = false;
  unit23OrderSummaryFile: File | null = null;
  unit23OrderSummaryError: string | null = null;
  unit23OrderSummarySummary: OrderSummarySummary = {};
  unit23OrderSummaryDownloadUrl = '';

  // Unit 23 Step 3: Sewing Actuals Upload
  unit23ActualsState: ActualsState = 'idle';
  unit23ActualsDragOver = false;
  unit23ActualsFile: File | null = null;
  unit23ActualsError: string | null = null;
  unit23ActualsSummary: FillActualsSummary = {};
  unit23ActualsDownloadUrl = '';

  // Unit 23 Unified Daily Actuals Upload
  unit23DirectWorkbookFile: File | null = null;
  unit23DirectSewingReportFile: File | null = null;
  unit23DirectDprReports: Unit23DprReportEntry[] = [this.createUnit23DprReportEntry()];
  unit23DirectActualsState: ActualsState = 'idle';
  unit23DirectActualsError: string | null = null;
  unit23DirectActualsSummary: FillActualsSummary = {};
  unit23DirectActualsDownloadUrl = '';
  unit23DprFilesProcessed = 0;

  onUnit23FileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) {
      this.unit23ErrorMessage = null;
      this.unit23SelectedFile = file;
    }
  }

  processUnit23(): void {
    if (!this.unit23SelectedFile) return;
    this.unit23ErrorMessage = null;
    this.unit23View = 'loading';

    this.linePlanService.processFileUnit23(this.unit23SelectedFile).subscribe({
      next: (res) => {
        this.unit23Summary = res.summary || {};
        this.unit23DownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.unit23PlanId = res.planId;
        this.unit23DownloadError = null;
        this.unit23ExcludedDates.clear();
        this.setUnit23AvailableDates(this.unit23Summary.availableDates || []);
        this.unit23View = 'result';
        this.refreshHistory();
      },
      error: (err: HttpErrorResponse) => {
        this.unit23ErrorMessage =
          err.error?.error || err.error?.message || err.message || 'Something went wrong while processing the Unit 23 file.';
        this.unit23View = 'upload';
      },
    });
  }

  resetUnit23(): void {
    this.unit23SelectedFile = null;
    this.unit23ErrorMessage = null;
    this.unit23Summary = {};
    this.unit23DownloadUrl = '';
    this.unit23PlanId = '';
    this.unit23DownloadError = null;
    this.unit23Downloading = false;
    this.unit23AvailableDates = [];
    this.unit23SelectedYear = '';
    this.unit23SelectedMonth = '';
    this.unit23ExcludedDates.clear();
    this.unit23View = 'upload';
    this.resetOrderSummaryUnit23();
    this.resetActualsUnit23();
  }

  setUnit23AvailableDates(dates: string[]): void {
    this.unit23AvailableDates = Array.from(new Set(dates)).sort();
    this.unit23SelectedYear = this.unit23Years[0] || '';
    this.unit23SelectedMonth = this.unit23Months[0] || '';
  }

  get unit23Years(): string[] {
    return Array.from(new Set(this.unit23AvailableDates.map((date) => date.slice(0, 4))));
  }

  get unit23Months(): string[] {
    if (!this.unit23SelectedYear) return [];
    return Array.from(
      new Set(
        this.unit23AvailableDates
          .filter((date) => date.startsWith(`${this.unit23SelectedYear}-`))
          .map((date) => date.slice(5, 7))
      )
    );
  }

  get unit23SelectedMonthTitle(): string {
    if (!this.unit23SelectedYear || !this.unit23SelectedMonth) return '';
    const date = new Date(Number(this.unit23SelectedYear), Number(this.unit23SelectedMonth) - 1, 1);
    return new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric' }).format(date);
  }

  get unit23ExcludedCount(): number {
    return this.unit23ExcludedDates.size;
  }

  get unit23CalendarWeeks(): CalendarCell[][] {
    if (!this.unit23SelectedYear || !this.unit23SelectedMonth) return [];

    const year = Number(this.unit23SelectedYear);
    const monthIndex = Number(this.unit23SelectedMonth) - 1;
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
    const firstWeekday = new Date(year, monthIndex, 1).getDay();
    const todayIso = this.toIsoDate(new Date());
    const available = new Set(this.unit23AvailableDates);
    const cells: CalendarCell[] = [];

    for (let i = 0; i < firstWeekday; i += 1) {
      cells.push({ day: null, iso: null, isWeekend: false, isToday: false, isAvailable: false });
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = new Date(year, monthIndex, day);
      const iso = this.toIsoDate(date);
      const weekday = date.getDay();
      cells.push({
        day,
        iso,
        isWeekend: weekday === 0 || weekday === 6,
        isToday: iso === todayIso,
        isAvailable: available.has(iso),
      });
    }

    const weeks: CalendarCell[][] = [];
    for (let i = 0; i < cells.length; i += 7) {
      weeks.push(cells.slice(i, i + 7));
    }
    return weeks;
  }

  onUnit23YearChange(): void {
    const months = this.unit23Months;
    if (!months.includes(this.unit23SelectedMonth)) {
      this.unit23SelectedMonth = months[0] || '';
    }
  }

  isUnit23Excluded(iso: string | null): boolean {
    return iso ? this.unit23ExcludedDates.has(iso) : false;
  }

  toggleUnit23Date(cell: CalendarCell): void {
    if (!cell.iso || !cell.isAvailable) return;
    if (this.unit23ExcludedDates.has(cell.iso)) {
      this.unit23ExcludedDates.delete(cell.iso);
    } else {
      this.unit23ExcludedDates.add(cell.iso);
    }
  }

  downloadGeneratedUnit23(): void {
    if (!this.unit23PlanId || this.unit23Downloading) return;
    this.unit23DownloadError = null;
    this.unit23Downloading = true;

    this.linePlanService
      .applyExclusions(this.unit23PlanId, Array.from(this.unit23ExcludedDates).sort())
      .subscribe({
        next: (res) => {
          this.unit23DownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
          this.unit23Downloading = false;
          window.location.href = this.unit23DownloadUrl;
        },
        error: (err: HttpErrorResponse) => {
          this.unit23DownloadError =
            err.error?.error || 'Something went wrong while preparing the download.';
          this.unit23Downloading = false;
        },
      });
  }

  // Unit 23 Step 2 Order Summary Handlers
  onUnit23OrderSummaryDragOver(event: DragEvent): void {
    event.preventDefault();
    this.unit23OrderSummaryDragOver = true;
  }

  onUnit23OrderSummaryDragLeave(): void {
    this.unit23OrderSummaryDragOver = false;
  }

  onUnit23OrderSummaryDrop(event: DragEvent): void {
    event.preventDefault();
    this.unit23OrderSummaryDragOver = false;
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setUnit23OrderSummaryFile(file);
  }

  onUnit23OrderSummaryFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.setUnit23OrderSummaryFile(file);
  }

  setUnit23OrderSummaryFile(file: File): void {
    this.unit23OrderSummaryError = null;
    this.unit23OrderSummaryFile = file;
  }

  clearUnit23OrderSummaryFile(): void {
    this.unit23OrderSummaryFile = null;
    this.unit23OrderSummaryError = null;
  }

  fillOrderSummaryUnit23(): void {
    if (!this.unit23OrderSummaryFile || !this.unit23PlanId) return;
    this.unit23OrderSummaryError = null;
    this.unit23OrderSummaryState = 'loading';

    const action$ = this.unit23ExcludedDates.size > 0
      ? this.linePlanService.applyExclusions(this.unit23PlanId, Array.from(this.unit23ExcludedDates).sort()).pipe(
          switchMap(() => this.linePlanService.fillOrderSummary(this.unit23PlanId, this.unit23OrderSummaryFile as File))
        )
      : this.linePlanService.fillOrderSummary(this.unit23PlanId, this.unit23OrderSummaryFile as File);

    action$.subscribe({
      next: (res) => {
        this.unit23OrderSummarySummary = res.summary || {};
        this.unit23OrderSummaryDownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.unit23DownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.unit23OrderSummaryState = 'done';
      },
      error: (err: HttpErrorResponse) => {
        this.unit23OrderSummaryError =
          err.error?.error || 'Something went wrong while processing the Order Sheet.';
        this.unit23OrderSummaryState = 'idle';
      },
    });
  }

  resetOrderSummaryUnit23(): void {
    this.unit23OrderSummaryState = 'idle';
    this.unit23OrderSummaryFile = null;
    this.unit23OrderSummaryError = null;
    this.unit23OrderSummarySummary = {};
    this.unit23OrderSummaryDownloadUrl = '';
  }

  // Unit 23 Step 3 Sewing Actuals Handlers
  onUnit23ActualsDragOver(event: DragEvent): void {
    event.preventDefault();
    this.unit23ActualsDragOver = true;
  }

  onUnit23ActualsDragLeave(): void {
    this.unit23ActualsDragOver = false;
  }

  onUnit23ActualsDrop(event: DragEvent): void {
    event.preventDefault();
    this.unit23ActualsDragOver = false;
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setUnit23ActualsFile(file);
  }

  onUnit23ActualsFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.setUnit23ActualsFile(file);
  }

  setUnit23ActualsFile(file: File): void {
    this.unit23ActualsError = null;
    this.unit23ActualsFile = file;
  }

  clearUnit23ActualsFile(): void {
    this.unit23ActualsFile = null;
    this.unit23ActualsError = null;
  }

  fillActualsUnit23(): void {
    if (!this.unit23ActualsFile || !this.unit23PlanId) return;
    this.unit23ActualsError = null;
    this.unit23ActualsState = 'loading';

    const action$ = this.unit23ExcludedDates.size > 0
      ? this.linePlanService.applyExclusions(this.unit23PlanId, Array.from(this.unit23ExcludedDates).sort()).pipe(
          switchMap(() => this.linePlanService.fillActuals(this.unit23PlanId, this.unit23ActualsFile as File))
        )
      : this.linePlanService.fillActuals(this.unit23PlanId, this.unit23ActualsFile as File);

    action$.subscribe({
      next: (res) => {
        this.unit23ActualsSummary = res.summary || {};
        this.unit23ActualsDownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.unit23ActualsState = 'done';
      },
      error: (err: HttpErrorResponse) => {
        this.unit23ActualsError =
          err.error?.error || 'Something went wrong while filling Sewing Actuals.';
        this.unit23ActualsState = 'idle';
      },
    });
  }

  resetActualsUnit23(): void {
    this.unit23ActualsState = 'idle';
    this.unit23ActualsFile = null;
    this.unit23ActualsError = null;
    this.unit23ActualsSummary = {};
    this.unit23ActualsDownloadUrl = '';
  }

  onUnit23DirectWorkbookFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.unit23DirectWorkbookFile = file;
  }


  onUnit23DirectSewingReportFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.unit23DirectSewingReportFile = file;
  }

  onUnit23DirectDprReportFileInputChange(event: Event, index: number): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.unit23DirectDprReports[index].file = file;
  }

  addUnit23DprReport(): void {
    this.unit23DirectDprReports.push(this.createUnit23DprReportEntry());
  }

  removeUnit23DprReport(index: number): void {
    this.unit23DirectDprReports.splice(index, 1);
  }

  clearUnit23DprReportFile(index: number): void {
    this.unit23DirectDprReports[index].file = null;
  }

  fillActualsDirectUnit23(): void {
    if (!this.unit23DirectWorkbookFile) return;
    this.unit23DirectActualsError = null;
    const selectedDprReports = this.unit23DirectDprReports.filter(
      (report): report is Unit23DprReportEntry & { file: File } => report.file !== null
    );
    if (selectedDprReports.some((report) => !report.reportDate)) {
      this.unit23DirectActualsError = 'Select a report date for every Wash & Pack DPR file.';
      return;
    }
    const selectedDates = selectedDprReports.map((report) => report.reportDate);
    if (new Set(selectedDates).size !== selectedDates.length) {
      this.unit23DirectActualsError = 'Each Wash & Pack DPR must have a different report date.';
      return;
    }

    const dprReports: DprReportUpload[] = selectedDprReports.map((report) => ({
      file: report.file,
      reportDate: report.reportDate,
    }));
    this.unit23DirectActualsState = 'loading';

    this.linePlanService
      .fillActualsDirectUnit23(
        this.unit23DirectWorkbookFile,
        this.unit23DirectSewingReportFile,
        dprReports
      )
      .subscribe({
        next: (res) => {
          if (!res.success) {
            this.unit23DirectActualsError = 'The update could not be completed. Please try again.';
            this.unit23DirectActualsState = 'idle';
            return;
          }
          this.unit23DirectActualsSummary = res.summary || {};
          this.unit23DirectActualsDownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
          this.unit23DprFilesProcessed = dprReports.length;
          this.unit23DirectActualsState = 'done';
        },
        error: (err: HttpErrorResponse) => {
          this.unit23DirectActualsError = err.error?.error || 'Failed to update Unit 23 daily actuals.';
          this.unit23DirectActualsState = 'idle';
        },
      });
  }

  resetDirectActualsUnit23(): void {
    this.unit23DirectActualsState = 'idle';
    this.unit23DirectSewingReportFile = null;
    this.unit23DirectDprReports = [this.createUnit23DprReportEntry()];
    this.unit23DirectActualsError = null;
    this.unit23DprFilesProcessed = 0;
  }

  private createUnit23DprReportEntry(): Unit23DprReportEntry {
    return { file: null, reportDate: new Date().toISOString().split('T')[0] };
  }


  // Step 2: upload a linewise performance report to fill Sewing Actuals
  // into the workbook generated in step 1.
  actualsState: ActualsState = 'idle';
  actualsDragOver = false;
  actualsFile: File | null = null;
  actualsError: string | null = null;
  actualsSummary: FillActualsSummary = {};
  actualsDownloadUrl = '';

  // Unified Direct Daily Actuals Upload Mode
  directWorkbookFile: File | null = null;
  directSewingReportFile: File | null = null;
  directWashPackReportFile: File | null = null;
  directActualsState: ActualsState = 'idle';
  directActualsError: string | null = null;
  directActualsSummary: FillActualsSummary = {};
  directActualsDownloadUrl = '';
  directUploadDragTarget: DirectUploadTarget | null = null;

  onDirectWorkbookFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.directWorkbookFile = file;
  }

  onDirectSewingReportFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.directSewingReportFile = file;
  }

  onDirectWashPackReportFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.directWashPackReportFile = file;
  }

  onDirectUploadDragOver(event: DragEvent, target: DirectUploadTarget): void {
    event.preventDefault();
    this.directUploadDragTarget = target;
  }

  onDirectUploadDragLeave(target: DirectUploadTarget): void {
    if (this.directUploadDragTarget === target) this.directUploadDragTarget = null;
  }

  onDirectUploadDrop(event: DragEvent, target: DirectUploadTarget, dprIndex?: number): void {
    event.preventDefault();
    this.directUploadDragTarget = null;
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;

    switch (target) {
      case 'unit1-workbook':
        this.directWorkbookFile = file;
        break;
      case 'unit1-sewing':
        this.directSewingReportFile = file;
        break;
      case 'unit1-wash-pack':
        this.directWashPackReportFile = file;
        break;
      case 'unit23-workbook':
        this.unit23DirectWorkbookFile = file;
        break;
      case 'unit23-sewing':
        this.unit23DirectSewingReportFile = file;
        break;
      case 'unit23-dpr':
        if (dprIndex !== undefined && this.unit23DirectDprReports[dprIndex]) {
          this.unit23DirectDprReports[dprIndex].file = file;
        }
        break;
    }
  }

  fillActualsDirect(): void {
    if (!this.directWorkbookFile) return;
    this.directActualsError = null;
    this.directActualsState = 'loading';

    this.linePlanService
      .fillActualsDirect(
        this.directWorkbookFile,
        this.directSewingReportFile,
        this.directWashPackReportFile
      )
      .subscribe({
        next: (res) => {
          if (!res.success) {
            this.directActualsError = 'The update could not be completed. Please try again.';
            this.directActualsState = 'idle';
            return;
          }
          this.directActualsSummary = res.summary || {};
          this.directActualsDownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
          this.directActualsState = 'done';
        },
        error: (err: HttpErrorResponse) => {
          this.directActualsError = err.error?.error || 'Failed to update daily actuals.';
          this.directActualsState = 'idle';
        },
      });
  }

  resetDirectActuals(): void {
    this.directActualsState = 'idle';
    this.directSewingReportFile = null;
    this.directWashPackReportFile = null;
    this.directActualsError = null;
  }

  dailyTab: 'sewing' | 'washpack' = 'sewing';

  // Standalone Daily Wash & Pack Actuals Upload Mode
  wpWorkbookFile: File | null = null;
  wpReportFile: File | null = null;
  wpActualsState: ActualsState = 'idle';
  wpActualsError: string | null = null;
  wpActualsSummary: FillActualsSummary = {};
  wpActualsDownloadUrl = '';

  jobs: JobRecord[] = [];
  dbAvailable = true;
  private healthTimer?: ReturnType<typeof setInterval>;

  constructor(private linePlanService: LinePlanService) {}

  ngOnInit(): void {
    this.refreshHistory();
    this.checkDatabaseHealth();
    this.healthTimer = setInterval(() => this.checkDatabaseHealth(), 15000);
  }

  ngOnDestroy(): void {
    if (this.healthTimer) clearInterval(this.healthTimer);
  }

  private checkDatabaseHealth(): void {
    this.linePlanService.health().subscribe({
      next: (res) => {
        const recovered = !this.dbAvailable && res.database;
        this.dbAvailable = res.database;
        if (recovered) this.refreshHistory();
      },
      error: () => (this.dbAvailable = false),
    });
  }

  onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.dragOver = true;
  }

  onDragLeave(): void {
    this.dragOver = false;
  }

  onDrop(event: DragEvent): void {
    event.preventDefault();
    this.dragOver = false;
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setFile(file);
  }

  onFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.setFile(file);
  }

  setFile(file: File): void {
    this.errorMessage = null;
    this.selectedFile = file;
  }

  clearFile(): void {
    this.selectedFile = null;
    this.errorMessage = null;
  }

  process(): void {
    if (!this.selectedFile) return;
    this.errorMessage = null;
    this.view = 'loading';

    this.linePlanService.processFile(this.selectedFile).subscribe({
      next: (res) => {
        this.summary = res.summary || {};
        this.downloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.planId = res.planId;
        this.downloadError = null;
        this.excludedDates.clear();
        this.setAvailableDates(this.summary.availableDates || []);
        this.view = 'result';
        this.refreshHistory();
      },
      error: (err: HttpErrorResponse) => {
        this.errorMessage =
          err.error?.error || 'Something went wrong while processing the file.';
        this.view = 'upload';
      },
    });
  }

  reset(): void {
    this.selectedFile = null;
    this.errorMessage = null;
    this.summary = {};
    this.downloadUrl = '';
    this.planId = '';
    this.downloadError = null;
    this.downloading = false;
    this.availableDates = [];
    this.selectedYear = '';
    this.selectedMonth = '';
    this.excludedDates.clear();
    this.view = 'upload';
    this.resetOrderSummary();
    this.resetActuals();
  }

  get years(): string[] {
    return Array.from(new Set(this.availableDates.map((date) => date.slice(0, 4))));
  }

  get months(): string[] {
    if (!this.selectedYear) return [];
    return Array.from(
      new Set(
        this.availableDates
          .filter((date) => date.startsWith(`${this.selectedYear}-`))
          .map((date) => date.slice(5, 7))
      )
    );
  }

  get selectedMonthTitle(): string {
    if (!this.selectedYear || !this.selectedMonth) return '';
    const date = new Date(Number(this.selectedYear), Number(this.selectedMonth) - 1, 1);
    return new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric' }).format(date);
  }

  get excludedCount(): number {
    return this.excludedDates.size;
  }

  get calendarWeeks(): CalendarCell[][] {
    if (!this.selectedYear || !this.selectedMonth) return [];

    const year = Number(this.selectedYear);
    const monthIndex = Number(this.selectedMonth) - 1;
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
    const firstWeekday = new Date(year, monthIndex, 1).getDay();
    const todayIso = this.toIsoDate(new Date());
    const available = new Set(this.availableDates);
    const cells: CalendarCell[] = [];

    for (let i = 0; i < firstWeekday; i += 1) {
      cells.push({ day: null, iso: null, isWeekend: false, isToday: false, isAvailable: false });
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
      const date = new Date(year, monthIndex, day);
      const iso = this.toIsoDate(date);
      const weekday = date.getDay();
      cells.push({
        day,
        iso,
        isWeekend: weekday === 0 || weekday === 6,
        isToday: iso === todayIso,
        isAvailable: available.has(iso),
      });
    }

    const weeks: CalendarCell[][] = [];
    for (let i = 0; i < cells.length; i += 7) {
      weeks.push(cells.slice(i, i + 7));
    }
    return weeks;
  }

  setAvailableDates(dates: string[]): void {
    this.availableDates = Array.from(new Set(dates)).sort();
    this.selectedYear = this.years[0] || '';
    this.selectedMonth = this.months[0] || '';
  }

  onYearChange(): void {
    const months = this.months;
    if (!months.includes(this.selectedMonth)) {
      this.selectedMonth = months[0] || '';
    }
  }

  monthName(month: string): string {
    return new Intl.DateTimeFormat('en-US', { month: 'long' }).format(
      new Date(2000, Number(month) - 1, 1)
    );
  }

  isExcluded(iso: string | null): boolean {
    return iso ? this.excludedDates.has(iso) : false;
  }

  toggleDate(cell: CalendarCell): void {
    if (!cell.iso || !cell.isAvailable) return;
    if (this.excludedDates.has(cell.iso)) {
      this.excludedDates.delete(cell.iso);
    } else {
      this.excludedDates.add(cell.iso);
    }
  }

  downloadGenerated(): void {
    if (!this.planId || this.downloading) return;
    this.downloadError = null;
    this.downloading = true;

    this.linePlanService
      .applyExclusions(this.planId, Array.from(this.excludedDates).sort())
      .subscribe({
        next: (res) => {
          this.downloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
          this.downloading = false;
          window.location.href = this.downloadUrl;
        },
        error: (err: HttpErrorResponse) => {
          this.downloadError =
            err.error?.error || 'Something went wrong while preparing the download.';
          this.downloading = false;
        },
      });
  }

  private toIsoDate(date: Date): string {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // ---- Step 1.5: Order Summary Upload ----

  onOrderSummaryDragOver(event: DragEvent): void {
    event.preventDefault();
    this.orderSummaryDragOver = true;
  }

  onOrderSummaryDragLeave(): void {
    this.orderSummaryDragOver = false;
  }

  onOrderSummaryDrop(event: DragEvent): void {
    event.preventDefault();
    this.orderSummaryDragOver = false;
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setOrderSummaryFile(file);
  }

  onOrderSummaryFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.setOrderSummaryFile(file);
  }

  setOrderSummaryFile(file: File): void {
    this.orderSummaryError = null;
    this.orderSummaryFile = file;
  }

  clearOrderSummaryFile(): void {
    this.orderSummaryFile = null;
    this.orderSummaryError = null;
  }

  fillOrderSummary(): void {
    if (!this.orderSummaryFile || !this.planId) return;
    this.orderSummaryError = null;
    this.orderSummaryState = 'loading';

    const action$ = this.excludedDates.size > 0
      ? this.linePlanService.applyExclusions(this.planId, Array.from(this.excludedDates).sort()).pipe(
          switchMap(() => this.linePlanService.fillOrderSummary(this.planId, this.orderSummaryFile as File))
        )
      : this.linePlanService.fillOrderSummary(this.planId, this.orderSummaryFile as File);

    action$.subscribe({
      next: (res) => {
        this.orderSummarySummary = res.summary || {};
        this.orderSummaryDownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.downloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.orderSummaryState = 'done';
      },
      error: (err: HttpErrorResponse) => {
        this.orderSummaryError =
          err.error?.error || 'Something went wrong while processing the Order Sheet.';
        this.orderSummaryState = 'idle';
      },
    });
  }

  resetOrderSummary(): void {
    this.orderSummaryState = 'idle';
    this.orderSummaryFile = null;
    this.orderSummaryError = null;
    this.orderSummarySummary = {};
    this.orderSummaryDownloadUrl = '';
  }

  // ---- Step 2: fill Sewing Actuals from a linewise performance report ----

  onActualsDragOver(event: DragEvent): void {
    event.preventDefault();
    this.actualsDragOver = true;
  }

  onActualsDragLeave(): void {
    this.actualsDragOver = false;
  }

  onActualsDrop(event: DragEvent): void {
    event.preventDefault();
    this.actualsDragOver = false;
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setActualsFile(file);
  }

  onActualsFileInputChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.setActualsFile(file);
  }

  setActualsFile(file: File): void {
    this.actualsError = null;
    this.actualsFile = file;
  }

  clearActualsFile(): void {
    this.actualsFile = null;
    this.actualsError = null;
  }

  fillActuals(): void {
    if (!this.actualsFile || !this.planId) return;
    this.actualsError = null;
    this.actualsState = 'loading';

    const action$ = this.excludedDates.size > 0
      ? this.linePlanService.applyExclusions(this.planId, Array.from(this.excludedDates).sort()).pipe(
          switchMap(() => this.linePlanService.fillActuals(this.planId, this.actualsFile as File))
        )
      : this.linePlanService.fillActuals(this.planId, this.actualsFile as File);

    action$.subscribe({
      next: (res) => {
        this.actualsSummary = res.summary || {};
        this.actualsDownloadUrl = this.linePlanService.getDownloadUrl(res.downloadUrl);
        this.actualsState = 'done';
      },
      error: (err: HttpErrorResponse) => {
        this.actualsError =
          err.error?.error || 'Something went wrong while filling Sewing Actuals.';
        this.actualsState = 'idle';
      },
    });
  }

  resetActuals(): void {
    this.actualsState = 'idle';
    this.actualsFile = null;
    this.actualsError = null;
    this.actualsSummary = {};
    this.actualsDownloadUrl = '';
  }



  refreshHistory(): void {
    this.linePlanService.listJobs().subscribe({
      next: (res) => (this.jobs = res.jobs || []),
      error: () => (this.jobs = []),
    });
  }

  statusClass(status: string): string {
    return `status-pill status-${status}`;
  }
}
