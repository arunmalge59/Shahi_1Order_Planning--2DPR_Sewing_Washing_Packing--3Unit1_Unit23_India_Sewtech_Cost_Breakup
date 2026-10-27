import { Injectable } from '@angular/core';

const API_BASE_URL = typeof window !== 'undefined' && window.location.hostname === 'localhost' && window.location.port === '4200' ? 'http://localhost:4000/planning' : '/planning';

@Injectable({
  providedIn: 'root',
})
export class HmNonDenimChartVariationService {
  constructor() {}

  async getAvailableMonths(previousWeekFile: File, currentWeekFile: File): Promise<string[]> {
    const formData = new FormData();
    formData.append('previousWeekFile', previousWeekFile);
    formData.append('currentWeekFile', currentWeekFile);

    const response = await fetch(`${API_BASE_URL}/api/hm/non-denim/chart-variation/months`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.error || `Server returned status ${response.status}`);
    }

    const body = await response.json();
    return Array.isArray(body.months) ? body.months : [];
  }

  async compareWeeklyCharts(previousWeekFile: File, currentWeekFile: File, startMonth?: string, endMonth?: string): Promise<Blob> {
    const formData = new FormData();
    formData.append('previousWeekFile', previousWeekFile);
    formData.append('currentWeekFile', currentWeekFile);
    if (startMonth) formData.append('startMonth', startMonth);
    if (endMonth) formData.append('endMonth', endMonth);

    const response = await fetch(`${API_BASE_URL}/api/hm/non-denim/chart-variation`, {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.error || `Server returned status ${response.status}`);
    }

    return response.blob();
  }
}
