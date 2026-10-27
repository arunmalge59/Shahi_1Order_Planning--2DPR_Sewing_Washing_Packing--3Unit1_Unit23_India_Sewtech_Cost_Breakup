import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { environment } from '../../../../../environments/environment';

@Injectable({ providedIn: 'root' })
export class ApiService {
  private base = '/cost/api';

  constructor(private http: HttpClient) {}

  uploadRaw(file: File, weekLabel: string, workingDays: number): Observable<any> {
    const form = new FormData();
    form.append('file', file);
    form.append('week_label', weekLabel);
    form.append('working_days', String(workingDays));
    return this.http.post(`${this.base}/batches/upload-raw`, form);
  }

  listBatches(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/batches`);
  }

  deleteBatch(batchId: number): Observable<any> {
    return this.http.delete(`${this.base}/batches/${batchId}`);
  }

  getReport(batchId: number): Observable<any> {
    return this.http.get(`${this.base}/batches/${batchId}/report`);
  }

  getRows(batchId: number): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/batches/${batchId}/rows`);
  }

  assignBuyer(batchId: number, payload: { brand: string; item?: string; buyer_group: string; save_to_master: boolean }): Observable<any> {
    return this.http.post(`${this.base}/batches/${batchId}/assign-buyer`, payload);
  }

  finalize(batchId: number): Observable<any> {
    return this.http.post(`${this.base}/batches/${batchId}/finalize`, {});
  }

  downloadModification(batchId: number): void {
    window.open(`${this.base}/batches/${batchId}/export-modification`, '_blank');
  }

  downloadFinal(batchId: number): void {
    window.open(`${this.base}/batches/${batchId}/final-export`, '_blank');
  }

  // Master data
  getBrandBuyerMap(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/master/brand-buyer`);
  }
  upsertBrandBuyer(payload: any): Observable<any> {
    return this.http.post(`${this.base}/master/brand-buyer`, payload);
  }
  deleteBrandBuyer(id: number): Observable<any> {
    return this.http.delete(`${this.base}/master/brand-buyer/${id}`);
  }
  getSamMaster(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/master/sam-master`);
  }
  upsertSamMaster(payload: any): Observable<any> {
    return this.http.post(`${this.base}/master/sam-master`, payload);
  }
}
