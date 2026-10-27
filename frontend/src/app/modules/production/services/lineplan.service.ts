import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from '../../../../environments/environment';

export interface ProcessSummary {
  styles?: number;
  rows?: number;
  dates?: number;
  filename?: string;
  availableDates?: string[];
}

export interface ProcessResponse {
  success: boolean;
  jobId: number | null;
  planId: string;
  downloadUrl: string;
  summary: ProcessSummary;
}

export interface FillActualsSummary {
  sheetsUpdated?: number;
  cellsFilled?: number;
  combinedSplitStyles?: string[];
  unmatchedStyles?: string[];
}

export interface FillActualsResponse {
  success: boolean;
  downloadUrl: string;
  summary: FillActualsSummary;
}

export interface DprReportUpload {
  file: File;
  reportDate: string;
}

export interface OrderSummarySummary {
  stylesMatched?: number;
  stylesNotFound?: number;
  matchedStyles?: string[];
  unmatchedStyles?: string[];
}

export interface OrderSummaryResponse {
  success: boolean;
  downloadUrl: string;
  summary: OrderSummarySummary;
}

export interface ApplyExclusionsResponse {
  success: boolean;
  downloadUrl: string;
  summary: {
    excludedDates?: string[];
    rowsRemoved?: number;
  };
}

export interface JobRecord {
  id: number;
  original_name: string;
  status: 'processing' | 'complete' | 'failed';
  styles_count: number | null;
  rows_count: number | null;
  dates_count: number | null;
  created_at: string;
  completed_at: string | null;
}

export interface HealthResponse {
  ok: boolean;
  database: boolean;
}

@Injectable({ providedIn: 'root' })
export class LinePlanService {
  private readonly baseUrl = environment.apiUrl.replace(/\/+$/, '');

  constructor(private http: HttpClient) {}

  processFile(file: File): Observable<ProcessResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<ProcessResponse>(`${this.baseUrl}/process`, formData);
  }

  fillOrderSummary(planId: string, file: File): Observable<OrderSummaryResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<OrderSummaryResponse>(`${this.baseUrl}/fill-order-summary/${planId}`, formData);
  }

  fillActuals(planId: string, file: File): Observable<FillActualsResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<FillActualsResponse>(`${this.baseUrl}/fill-actuals/${planId}`, formData);
  }

  fillActualsDirect(
    workbook: File,
    sewingReport?: File | null,
    washPackReport?: File | null
  ): Observable<FillActualsResponse> {
    const formData = new FormData();
    formData.append('workbook', workbook);
    if (sewingReport) {
      formData.append('sewingReport', sewingReport);
    }
    if (washPackReport) {
      formData.append('washPackReport', washPackReport);
    }
    return this.http.post<FillActualsResponse>(`${this.baseUrl}/fill-actuals-direct`, formData);
  }

  fillWashPack(planId: string, file: File): Observable<FillActualsResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<FillActualsResponse>(`${this.baseUrl}/fill-wash-pack/${planId}`, formData);
  }

  fillWashPackDirect(workbook: File, report?: File | null): Observable<FillActualsResponse> {
    const formData = new FormData();
    formData.append('workbook', workbook);
    if (report) {
      formData.append('report', report);
    }
    return this.http.post<FillActualsResponse>(`${this.baseUrl}/fill-wash-pack-direct`, formData);
  }

  applyExclusions(planId: string, excludedDates: string[]): Observable<ApplyExclusionsResponse> {
    return this.http.post<ApplyExclusionsResponse>(`${this.baseUrl}/apply-exclusions/${planId}`, {
      excludedDates,
    });
  }

  processFileUnit23(file: File): Observable<ProcessResponse> {
    const formData = new FormData();
    formData.append('file', file);
    return this.http.post<ProcessResponse>(`${this.baseUrl}/unit23/process`, formData);
  }

  fillActualsDirectUnit23(
    workbook: File,
    sewingReport?: File | null,
    dprReports: DprReportUpload[] = []
  ): Observable<FillActualsResponse> {
    const formData = new FormData();
    formData.append('workbook', workbook);
    if (sewingReport) {
      formData.append('sewingReport', sewingReport);
    }
    dprReports.forEach((report, index) => {
      formData.append(`dprReport_${index}`, report.file);
      formData.append(`dprDate_${index}`, report.reportDate);
    });
    return this.http.post<FillActualsResponse>(`${this.baseUrl}/unit23/fill-actuals-direct`, formData);
  }

  getDownloadUrl(relativeUrl: string): string {
    if (!relativeUrl) return '';
    if (relativeUrl.startsWith('http://') || relativeUrl.startsWith('https://')) {
      return relativeUrl;
    }
    const cleanRelative = relativeUrl.startsWith('/api/')
      ? relativeUrl.substring(4)
      : relativeUrl.startsWith('/')
      ? relativeUrl
      : `/${relativeUrl}`;
    return `${this.baseUrl}${cleanRelative}`;
  }

  listJobs(): Observable<{ jobs: JobRecord[] }> {
    return this.http.get<{ jobs: JobRecord[] }>(`${this.baseUrl}/jobs`);
  }

  health(): Observable<HealthResponse> {
    return this.http.get<HealthResponse>(`${this.baseUrl}/health`);
  }
}
