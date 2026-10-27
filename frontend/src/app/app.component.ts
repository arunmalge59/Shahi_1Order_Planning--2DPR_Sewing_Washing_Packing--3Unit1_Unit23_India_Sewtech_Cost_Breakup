import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { NavigationEnd, Router, RouterOutlet } from '@angular/router';
import { filter } from 'rxjs';
import { AppNavigationComponent } from './app-navigation.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, AppNavigationComponent, RouterOutlet],
  template: `
    <div class="unified-shell" [class.rail-layout]="sidebarCollapsed" [class.production-layout]="isProductionRoute">
      <app-navigation
        [collapsed]="sidebarCollapsed"
        [mobileOpen]="mobileSidebarOpen"
        (collapseToggle)="toggleSidebar()"
        (mobileClose)="closeMobileSidebar()"
      ></app-navigation>

      <button
        *ngIf="mobileSidebarOpen"
        class="sidebar-backdrop"
        type="button"
        aria-label="Close navigation"
        (click)="closeMobileSidebar()"
      ></button>

      <div class="shell-main">
        <header class="shell-topbar" [class.production-menu-bar]="isProductionRoute">
          <button
            class="shell-menu-toggle"
            type="button"
            [attr.aria-label]="mobileSidebarOpen ? 'Close navigation' : 'Open navigation'"
            (click)="toggleSidebar()"
          >
            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16"></path></svg>
          </button>
          <div *ngIf="!isProductionRoute">
            <p class="shell-eyebrow">SHAHI · UNIFIED OPERATIONS</p>
            <h1>{{ sectionTitle }}</h1>
          </div>
        </header>
        <main
          class="workflow-content"
          [class.production-content]="isProductionRoute"
          [class.cost-content]="isCostRoute"
        >
          <router-outlet></router-outlet>
        </main>
      </div>
    </div>
  `
})
export class AppComponent {
  sidebarCollapsed = false;
  mobileSidebarOpen = false;
  private currentUrl: string;

  constructor(private readonly router: Router) {
    this.currentUrl = window.location.pathname === '/' ? '/production' : window.location.pathname;
    this.router.events
      .pipe(filter((event): event is NavigationEnd => event instanceof NavigationEnd))
      .subscribe((event) => {
        this.currentUrl = event.urlAfterRedirects;
        this.closeMobileSidebar();
      });
  }

  get sectionTitle(): string {
    const url = this.currentUrl;
    if (url.startsWith('/cost')) return 'India Cost Breakup';
    if (url.startsWith('/planning')) return 'Planning and Order';
    return 'Daily DPR';
  }

  get isProductionRoute(): boolean {
    return this.currentUrl.startsWith('/production');
  }

  get isCostRoute(): boolean {
    return this.currentUrl.startsWith('/cost');
  }

  toggleSidebar(): void {
    if (window.matchMedia('(max-width: 767px)').matches) {
      this.mobileSidebarOpen = !this.mobileSidebarOpen;
      return;
    }
    this.sidebarCollapsed = !this.sidebarCollapsed;
  }

  closeMobileSidebar(): void {
    this.mobileSidebarOpen = false;
  }
}
