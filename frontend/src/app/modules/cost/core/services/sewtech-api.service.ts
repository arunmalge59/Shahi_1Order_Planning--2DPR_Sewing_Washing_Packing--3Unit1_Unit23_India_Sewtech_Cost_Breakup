import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { environment } from '../../../../../environments/environment';

@Injectable({ providedIn: 'root' })
export class SewtechApiService {
  private base = '/cost/api/sewtech';

  constructor(private http: HttpClient) {}

  uploadRaw(file: File): Observable<any> {
    const form = new FormData();
    form.append('file', file);
    return this.http.post(`${this.base}/upload-raw`, form);
  }

  processBatch(batchId: number, payload: { start_date: string; end_date: string; working_days?: number }): Observable<any> {
    return this.http.post(`${this.base}/${batchId}/process`, payload);
  }

  getReport(batchId: number): Observable<any> {
    return this.http.get(`${this.base}/${batchId}/report`);
  }

  getRows(batchId: number, limit = 300, offset = 0): Observable<any> {
    return this.http.get(`${this.base}/${batchId}/rows?limit=${limit}&offset=${offset}`);
  }

  assignBuyer(batchId: number, payload: { raw_buyer: string; buyer_group?: string; new_group?: string; save_mapping: boolean }): Observable<any> {
    return this.http.post(`${this.base}/${batchId}/assign-buyer`, payload);
  }

  decideHp(batchId: number, includeHp: boolean): Observable<any> {
    return this.http.post(`${this.base}/${batchId}/hp-decision`, { include_hp: includeHp });
  }

  updateWorkingDays(batchId: number, workingDays: number): Observable<any> {
    return this.http.post(`${this.base}/${batchId}/working-days`, { working_days: workingDays });
  }

  finalize(batchId: number): Observable<any> {
    return this.http.post(`${this.base}/${batchId}/finalize`, {});
  }

  uploadModified(batchId: number, file: File): Observable<any> {
    const form = new FormData();
    form.append('file', file);
    return this.http.post(`${this.base}/${batchId}/upload-modified`, form);
  }

  downloadMerged(batchId: number): void {
    window.open(`${this.base}/${batchId}/export-merged`, '_blank');
  }

  downloadFinal(batchId: number): void {
    window.open(`${this.base}/${batchId}/final-export`, '_blank');
  }

  listBatches(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}`);
  }

  deleteBatch(batchId: number): Observable<any> {
    return this.http.delete(`${this.base}/${batchId}`);
  }

  // Master data
  getBuyerGroups(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/master/buyer-groups`);
  }

  addBuyerGroup(name: string): Observable<any> {
    return this.http.post(`${this.base}/master/buyer-groups`, { name });
  }

  getBuyerMap(): Observable<any[]> {
    return this.http.get<any[]>(`${this.base}/master/buyer-map`);
  }

  upsertBuyerMap(payload: { raw_buyer: string; buyer_group_id?: number; buyer_group?: string }): Observable<any> {
    return this.http.post(`${this.base}/master/buyer-map`, payload);
  }

  deleteBuyerMap(id: number): Observable<any> {
    return this.http.delete(`${this.base}/master/buyer-map/${id}`);
  }
}
