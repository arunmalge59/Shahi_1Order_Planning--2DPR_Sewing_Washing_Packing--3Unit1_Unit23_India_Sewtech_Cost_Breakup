import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { SewtechApiService } from '../core/services/sewtech-api.service';

@Component({
  selector: 'app-sewtech-master',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="container">
      <!-- Buyer Mapping Card -->
      <div class="card">
        <div style="display: flex; align-items: flex-start; gap: 14px; margin-bottom: 16px;">
          <div style="width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #f7ebfa 0%, #e6f6f7 100%); display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">
            🏷️
          </div>
          <div>
            <h2 style="margin: 0 0 4px 0;">Sewtech Raw Buyer → Buyer Group Mapping</h2>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              Raw buyers discovered in daily sheets are mapped to standard buyer groups (e.g. <code>BD Design → Job Work</code>). Mappings saved here apply automatically across all future weeks.
            </p>
          </div>
        </div>

        <!-- Add Mapping Form -->
        <div style="display: flex; gap: 12px; margin: 20px 0; flex-wrap: wrap; align-items: flex-end;">
          <div>
            <label style="display: block; margin-bottom: 4px;">Raw Buyer Name</label>
            <input
              type="text"
              placeholder="e.g. Texas Vintage"
              [(ngModel)]="newMap.raw_buyer"
              style="min-width: 180px;"
            />
          </div>

          <div>
            <label style="display: block; margin-bottom: 4px;">Buyer Group</label>
            <select [(ngModel)]="newMap.buyer_group" style="min-width: 180px;">
              <option value="">-- Choose Group --</option>
              <option *ngFor="let g of buyerGroups" [value]="g.name">{{ g.name }}</option>
            </select>
          </div>

          <button
            class="btn"
            [disabled]="!newMap.raw_buyer.trim() || !newMap.buyer_group"
            (click)="addMapping()"
          >
            Add / Update Mapping
          </button>
        </div>

        <div style="border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
          <table>
            <thead>
              <tr>
                <th>Raw Buyer in Excel</th>
                <th>Normalized Key</th>
                <th>Assigned Buyer Group</th>
                <th>Created</th>
                <th style="text-align: right;">Action</th>
              </tr>
            </thead>
            <tbody>
              <tr *ngFor="let m of buyerMap">
                <td><strong>{{ m.raw_buyer_display }}</strong></td>
                <td><code>{{ m.raw_buyer_norm }}</code></td>
                <td><span class="badge teal">{{ m.buyer_group }}</span></td>
                <td>{{ m.created_at | date:'dd-MMM-yyyy' }}</td>
                <td style="text-align: right;">
                  <button class="btn-delete" (click)="deleteMapping(m.id)">
                    Delete
                  </button>
                </td>
              </tr>
              <tr *ngIf="buyerMap.length === 0">
                <td colspan="5" style="text-align: center; color: #94a3b8; padding: 24px;">No buyer mappings saved yet.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Buyer Groups Card -->
      <div class="card">
        <div style="display: flex; align-items: flex-start; gap: 14px; margin-bottom: 16px;">
          <div style="width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #f7ebfa 0%, #e6f6f7 100%); display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">
            🏢
          </div>
          <div>
            <h2 style="margin: 0 0 4px 0;">Sewtech Buyer Groups Master</h2>
            <p style="margin: 0; color: #64748b; font-size: 13px;">
              The summary report aggregates rows according to these buyer groups. Add custom groups as needed.
            </p>
          </div>
        </div>

        <div style="display: flex; gap: 12px; margin: 18px 0; align-items: center; flex-wrap: wrap;">
          <input
            type="text"
            placeholder="New group name (e.g. Job Work)"
            [(ngModel)]="newGroupName"
            style="min-width: 220px;"
          />
          <button class="btn" [disabled]="!newGroupName.trim()" (click)="addGroup()">Add Group</button>
        </div>

        <div style="display: flex; gap: 10px; flex-wrap: wrap; margin-top: 12px;">
          <span
            *ngFor="let g of buyerGroups"
            class="badge purple"
            style="font-size: 13px; padding: 6px 14px;"
          >
            {{ g.name }}
          </span>
        </div>
      </div>
    </div>
  `,
})
export class SewtechMasterComponent implements OnInit {
  buyerMap: any[] = [];
  buyerGroups: any[] = [];

  newMap = { raw_buyer: '', buyer_group: '' };
  newGroupName = '';

  constructor(private api: SewtechApiService) {}

  ngOnInit() {
    this.reload();
  }

  reload() {
    this.api.getBuyerMap().subscribe((res) => (this.buyerMap = res || []));
    this.api.getBuyerGroups().subscribe((res) => (this.buyerGroups = res || []));
  }

  addMapping() {
    if (!this.newMap.raw_buyer?.trim() || !this.newMap.buyer_group) return;
    this.api.upsertBuyerMap({
      raw_buyer: this.newMap.raw_buyer.trim(),
      buyer_group: this.newMap.buyer_group,
    }).subscribe(() => {
      this.newMap = { raw_buyer: '', buyer_group: '' };
      this.reload();
    });
  }

  deleteMapping(id: number) {
    if (!confirm('Are you sure you want to delete this mapping?')) return;
    this.api.deleteBuyerMap(id).subscribe(() => {
      this.reload();
    });
  }

  addGroup() {
    if (!this.newGroupName?.trim()) return;
    this.api.addBuyerGroup(this.newGroupName.trim()).subscribe(() => {
      this.newGroupName = '';
      this.reload();
    });
  }
}
