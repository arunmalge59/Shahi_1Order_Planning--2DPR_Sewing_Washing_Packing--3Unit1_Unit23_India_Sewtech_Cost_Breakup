import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { ApiService } from '../core/services/api.service';

@Component({
  selector: 'app-master-data',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="container">
      <!-- Brand Buyer Map -->
      <div class="card">
        <div style="display: flex; align-items: flex-start; gap: 14px; margin-bottom: 16px;">
          <div style="width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #e6f6f7 0%, #f7ebfa 100%); display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">
            🏷️
          </div>
          <div>
            <h2 style="margin: 0 0 4px 0;">India Factory — Brand → Buyer Group Mapping</h2>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              Define mapping rules from raw Brand codes to designated Buyer Groups. Future weekly uploads resolve automatically without manual intervention.
            </p>
          </div>
        </div>

        <div style="display: flex; gap: 12px; margin: 20px 0; flex-wrap: wrap; align-items: center;">
          <input
            type="text"
            placeholder="Brand (e.g. KOHLS)"
            [(ngModel)]="newBrand.brand"
            style="min-width: 160px;"
          />
          <input
            type="text"
            placeholder="Buyer Group (e.g. Kohls)"
            [(ngModel)]="newBrand.buyer_group"
            style="min-width: 180px;"
          />
          <label style="display: inline-flex; align-items: center; gap: 6px; cursor: pointer;">
            <input type="checkbox" [(ngModel)]="newBrand.is_denim_split" />
            <span>Splits by item (e.g. denim / non-denim)</span>
          </label>
          <button
            class="btn"
            [disabled]="!newBrand.brand.trim() || !newBrand.buyer_group.trim()"
            (click)="addBrand()"
          >
            Add / Update Mapping
          </button>
        </div>

        <div style="border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
          <table>
            <thead>
              <tr>
                <th>Brand</th>
                <th>Assigned Buyer Group</th>
                <th>Splits by Item?</th>
                <th style="text-align: right;">Action</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let m of brandMap">
                <td><strong>{{ m.brand }}</strong></td>
                <td><span class="badge teal">{{ m.buyer_group }}</span></td>
                <td>
                  <span class="badge" [ngClass]="{'purple': m.is_denim_split}">
                    {{ m.is_denim_split ? 'Yes' : 'No' }}
                  </span>
                </td>
                <td style="text-align: right;">
                  <button class="btn-delete" (click)="deleteBrand(m.id)">Delete</button>
                </td>
              </tr>
              <tr *ngIf="brandMap.length === 0">
                <td colspan="4" style="text-align: center; color: #94a3b8; padding: 24px;">
                  No brand mappings configured yet.
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- SAM Master -->
      <div class="card">
        <div style="display: flex; align-items: flex-start; gap: 14px; margin-bottom: 16px;">
          <div style="width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #e6f6f7 0%, #f7ebfa 100%); display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">
            ⏱️
          </div>
          <div>
            <h2 style="margin: 0 0 4px 0;">Standard SAM Per Piece Master</h2>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              Used as a fallback standard for new styles where raw SAM = 0 to accurately compute operator & machine hours.
            </p>
          </div>
        </div>

        <div style="display: flex; gap: 12px; margin: 20px 0; flex-wrap: wrap; align-items: center;">
          <input
            type="text"
            placeholder="Item / Style Code"
            [(ngModel)]="newSam.item"
            style="min-width: 180px;"
          />
          <input
            type="number"
            placeholder="Standard SAM per piece"
            [(ngModel)]="newSam.standard_sam_per_piece"
            style="min-width: 180px;"
          />
          <button
            class="btn"
            [disabled]="!newSam.item.trim() || newSam.standard_sam_per_piece == null"
            (click)="addSam()"
          >
            Add / Update SAM
          </button>
        </div>

        <div style="border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
          <table>
            <thead>
              <tr>
                <th>Item / Style Code</th>
                <th>Standard SAM / Piece</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let s of samMaster">
                <td><strong>{{ s.item }}</strong></td>
                <td>{{ s.standard_sam_per_piece }}</td>
              </tr>
              <tr *ngIf="samMaster.length === 0">
                <td colspan="2" style="text-align: center; color: #94a3b8; padding: 24px;">
                  No SAM master records saved yet.
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `,
})
export class MasterDataComponent implements OnInit {
  brandMap: any[] = [];
  samMaster: any[] = [];
  newBrand = { brand: '', buyer_group: '', is_denim_split: false };
  newSam = { item: '', standard_sam_per_piece: null as number | null };

  constructor(private api: ApiService) {}

  ngOnInit() {
    this.reload();
  }

  reload() {
    this.api.getBrandBuyerMap().subscribe((r) => (this.brandMap = r || []));
    this.api.getSamMaster().subscribe((r) => (this.samMaster = r || []));
  }

  addBrand() {
    if (!this.newBrand.brand?.trim() || !this.newBrand.buyer_group?.trim()) return;
    this.api.upsertBrandBuyer(this.newBrand).subscribe(() => {
      this.newBrand = { brand: '', buyer_group: '', is_denim_split: false };
      this.reload();
    });
  }

  deleteBrand(id: number) {
    if (!confirm('Are you sure you want to delete this brand mapping?')) return;
    this.api.deleteBrandBuyer(id).subscribe(() => this.reload());
  }

  addSam() {
    if (!this.newSam.item?.trim() || this.newSam.standard_sam_per_piece == null) return;
    this.api.upsertSamMaster(this.newSam).subscribe(() => {
      this.newSam = { item: '', standard_sam_per_piece: null };
      this.reload();
    });
  }
}
