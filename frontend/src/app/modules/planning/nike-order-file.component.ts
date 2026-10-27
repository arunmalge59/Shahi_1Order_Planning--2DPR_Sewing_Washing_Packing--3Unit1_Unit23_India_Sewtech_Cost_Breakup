import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';

const API_BASE_URL = typeof window !== 'undefined' && window.location.hostname === 'localhost' && window.location.port === '4200' ? 'http://localhost:4000/planning' : '/planning';

@Component({
  selector: 'app-nike-order-file',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './nike-order-file.component.html',
  styleUrl: './nike-order-file.component.css',
})
export class NikeOrderFileComponent {
  file: File | null = null;
  weekCode = 'W36';
  loading = false;
  errorMessage = '';
  generatedFile: Blob | null = null;
  

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.file = input.files?.[0] || null;
    this.generatedFile = null;
    this.errorMessage = '';
  }

  async createOrderFile(): Promise<void> {
    if (!this.file) return;
    this.loading = true;
    this.errorMessage = '';
    this.generatedFile = null;
    const formData = new FormData();
    formData.append('file', this.file);
    formData.append('weekCode', this.weekCode.trim() || 'W36');

    try {
      const response = await fetch(`${API_BASE_URL}/api/order-file/generate`, { method: 'POST', body: formData });
      if (!response.ok) {
        const body = await response.json().catch(() => ({}));
        throw new Error(body.error || `Server returned status ${response.status}`);
      }
      this.generatedFile = await response.blob();
    } catch (error: any) {
      this.errorMessage = error.message || 'Could not create the order file.';
    } finally {
      this.loading = false;
    }
  }

  downloadOrderFile(): void {
    if (!this.generatedFile) return;
    const link = document.createElement('a');
    link.href = URL.createObjectURL(this.generatedFile);
    link.download = 'nike-order-file.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  }
}
