import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { NavigationEnd, Router, RouterLink, RouterLinkActive, RouterOutlet } from '@angular/router';
import { filter } from 'rxjs';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet, RouterLink, RouterLinkActive],
  template: `
    <div class="app-layout">
      <!-- Top Global Header -->
      <header class="top-header">
        <div class="header-left">
          <div class="shahi-logo-group">
            <span class="shahi-brand-text">shahi</span>
            <span class="automation-title">AUTOMATION</span>
          </div>
          <div class="factory-context-pill">
            <span class="dot" [ngClass]="currentModule"></span>
            <span>{{ currentModule === 'india' ? 'India Factory' : 'Sewtech Factory' }}</span>
          </div>
        </div>

        <div class="header-right">
          <div class="system-status">
            <span class="status-indicator"></span>
            <span>Production Ready v2.0</span>
          </div>
        </div>
      </header>

      <div class="app-body">
        <!-- Left Sidebar Navigation -->
        <aside class="sidebar">
          <div class="sidebar-section-title">FACTORIES & WORKFLOWS</div>

          <!-- India Factory Accordion -->
          <div class="factory-group">
            <button
              class="factory-header-btn"
              [class.active-factory]="currentModule === 'india'"
              (click)="toggleIndia()"
              type="button"
            >
              <div class="factory-header-title">
                <span class="factory-icon">🏭</span>
                <span>India Factory</span>
              </div>
              <span class="arrow-icon" [class.open]="isIndiaOpen">▾</span>
            </button>

            <div class="submenu" *ngIf="isIndiaOpen">
              <a
                routerLink="/india/upload"
                routerLinkActive="active"
                [routerLinkActiveOptions]="{ exact: true }"
                class="sidebar-link"
              >
                <span class="menu-bullet"></span>
                <span>Upload Raw</span>
              </a>
              <a
                routerLink="/india/batches"
                routerLinkActive="active"
                class="sidebar-link"
              >
                <span class="menu-bullet"></span>
                <span>Weekly Batches</span>
              </a>
              <a
                routerLink="/india/master-data"
                routerLinkActive="active"
                class="sidebar-link"
              >
                <span class="menu-bullet"></span>
                <span>Master Data</span>
              </a>
            </div>
          </div>

          <!-- Sewtech Factory Accordion -->
          <div class="factory-group">
            <button
              class="factory-header-btn"
              [class.active-factory]="currentModule === 'sewtech'"
              (click)="toggleSewtech()"
              type="button"
            >
              <div class="factory-header-title">
                <span class="factory-icon">🏭</span>
                <span>Sewtech Factory</span>
              </div>
              <span class="arrow-icon" [class.open]="isSewtechOpen">▾</span>
            </button>

            <div class="submenu" *ngIf="isSewtechOpen">
              <a
                routerLink="/sewtech/upload"
                routerLinkActive="active"
                [routerLinkActiveOptions]="{ exact: true }"
                class="sidebar-link"
              >
                <span class="menu-bullet"></span>
                <span>Upload Raw</span>
              </a>
              <a
                routerLink="/sewtech/batches"
                routerLinkActive="active"
                class="sidebar-link"
              >
                <span class="menu-bullet"></span>
                <span>Weekly Batches</span>
              </a>
              <a
                routerLink="/sewtech/master-data"
                routerLinkActive="active"
                class="sidebar-link"
              >
                <span class="menu-bullet"></span>
                <span>Buyer Mapping & Master</span>
              </a>
            </div>
          </div>

          <div class="sidebar-footer">
            <div class="shahi-brand-tagline">
              Shahi Exports Pvt. Ltd.
            </div>
          </div>
        </aside>

        <!-- Main Content Area -->
        <main class="main-content">
          <router-outlet></router-outlet>
        </main>
      </div>
    </div>
  `,
  styles: [`
    .app-layout {
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      background-color: var(--bg-main);
    }

    /* Top Global Header */
    .top-header {
      height: 60px;
      background: #ffffff;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 0 24px;
      position: sticky;
      top: 0;
      z-index: 100;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03);
    }

    .header-left {
      display: flex;
      align-items: center;
      gap: 20px;
    }

    .shahi-logo-group {
      display: flex;
      align-items: baseline;
      gap: 8px;
    }

    .shahi-brand-text {
      font-size: 24px;
      font-weight: 800;
      letter-spacing: -0.5px;
      font-family: 'Plus Jakarta Sans', sans-serif;
      background: linear-gradient(135deg, #0e8b94 0%, #1e6292 45%, #6a2578 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      line-height: 1;
    }

    .automation-title {
      font-size: 13px;
      font-weight: 700;
      letter-spacing: 1.5px;
      color: #334155;
    }

    .factory-context-pill {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 12px;
      background: #f1f5f9;
      border: 1px solid #e2e8f0;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      color: #334155;
    }

    .factory-context-pill .dot {
      width: 7px;
      height: 7px;
      border-radius: 50%;
    }

    .factory-context-pill .dot.india {
      background: #0e8b94;
      box-shadow: 0 0 6px rgba(14, 139, 148, 0.6);
    }

    .factory-context-pill .dot.sewtech {
      background: #6a2578;
      box-shadow: 0 0 6px rgba(106, 37, 120, 0.6);
    }

    .header-right {
      display: flex;
      align-items: center;
      gap: 16px;
    }

    .system-status {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 12.5px;
      font-weight: 500;
      color: #64748b;
      background: #f8fafc;
      padding: 6px 12px;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
    }

    .status-indicator {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #10b981;
      box-shadow: 0 0 6px rgba(16, 185, 129, 0.5);
    }

    /* App Body */
    .app-body {
      display: flex;
      flex: 1;
    }

    /* Left Sidebar */
    .sidebar {
      width: 260px;
      background: #ffffff;
      border-right: 1px solid var(--border-color);
      display: flex;
      flex-direction: column;
      padding: 20px 12px;
      position: sticky;
      top: 60px;
      height: calc(100vh - 60px);
      overflow-y: auto;
      flex-shrink: 0;
    }

    .sidebar-section-title {
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 0.08em;
      color: #94a3b8;
      padding: 0 12px 10px 12px;
    }

    .factory-group {
      margin-bottom: 12px;
    }

    .factory-header-btn {
      width: 100%;
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 14px;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 8px;
      font-family: inherit;
      font-size: 13.5px;
      font-weight: 600;
      color: #334155;
      cursor: pointer;
      transition: all 0.15s ease;
      text-align: left;
    }

    .factory-header-btn:hover {
      background: #f8fafc;
      color: #0f172a;
    }

    .factory-header-btn.active-factory {
      background: #f1f5f9;
      color: #0f172a;
      border-color: #e2e8f0;
    }

    .factory-header-title {
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .factory-icon {
      font-size: 16px;
    }

    .arrow-icon {
      font-size: 12px;
      color: #94a3b8;
      transition: transform 0.2s ease;
      display: inline-block;
    }

    .arrow-icon.open {
      transform: rotate(0deg);
    }

    .arrow-icon:not(.open) {
      transform: rotate(-90deg);
    }

    .submenu {
      display: flex;
      flex-direction: column;
      margin-top: 4px;
      padding-left: 8px;
    }

    .sidebar-link {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 8px 12px 8px 24px;
      margin: 2px 0;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 500;
      color: #64748b;
      text-decoration: none;
      transition: all 0.15s ease;
      position: relative;
    }

    .sidebar-link:hover {
      background: #f8fafc;
      color: #0f172a;
    }

    .sidebar-link.active {
      background: linear-gradient(135deg, rgba(14, 139, 148, 0.1) 0%, rgba(106, 37, 120, 0.08) 100%);
      color: #0e8b94;
      font-weight: 700;
      border-left: 3px solid #0e8b94;
    }

    .menu-bullet {
      width: 5px;
      height: 5px;
      border-radius: 50%;
      background: #cbd5e1;
      transition: background 0.15s ease;
    }

    .sidebar-link.active .menu-bullet {
      background: #0e8b94;
      box-shadow: 0 0 4px rgba(14, 139, 148, 0.5);
    }

    .sidebar-footer {
      margin-top: auto;
      padding: 16px 12px 0 12px;
      border-top: 1px solid #f1f5f9;
    }

    .shahi-brand-tagline {
      font-size: 11px;
      color: #94a3b8;
      font-weight: 500;
    }

    /* Main Content */
    .main-content {
      flex: 1;
      min-width: 0;
      min-height: calc(100vh - 60px);
      background: var(--bg-main);
      padding-bottom: 40px;
    }
  `],
})
export class AppComponent {
  currentModule: 'india' | 'sewtech' = 'india';
  isIndiaOpen = true;
  isSewtechOpen = false;

  constructor(private router: Router) {
    this.router.events
      .pipe(filter((e): e is NavigationEnd => e instanceof NavigationEnd))
      .subscribe((event) => {
        if (event.urlAfterRedirects.includes('/sewtech')) {
          this.currentModule = 'sewtech';
          this.isSewtechOpen = true;
        } else {
          this.currentModule = 'india';
          this.isIndiaOpen = true;
        }
      });
  }

  toggleIndia() {
    this.isIndiaOpen = !this.isIndiaOpen;
    if (this.isIndiaOpen && this.currentModule !== 'india') {
      this.router.navigate(['/india/upload']);
    }
  }

  toggleSewtech() {
    this.isSewtechOpen = !this.isSewtechOpen;
    if (this.isSewtechOpen && this.currentModule !== 'sewtech') {
      this.router.navigate(['/sewtech/upload']);
    }
  }
}
