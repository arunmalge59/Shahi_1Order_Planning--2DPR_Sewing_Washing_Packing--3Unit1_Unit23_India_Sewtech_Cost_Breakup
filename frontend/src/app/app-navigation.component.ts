import { CommonModule } from '@angular/common';
import { Component, EventEmitter, Input, OnDestroy, OnInit, Output } from '@angular/core';
import { IsActiveMatchOptions, NavigationEnd, Router, RouterLink } from '@angular/router';
import { filter, Subscription } from 'rxjs';

type ProjectId = 'production' | 'cost' | 'planning';

interface WorkflowLink {
  section: string;
  label: string;
  route: string;
  queryParams?: Record<string, string>;
}

interface ProjectGroup {
  id: ProjectId;
  number: string;
  title: string;
  subtitle: string;
  icon: 'production' | 'cost' | 'planning';
  route: string;
  links: WorkflowLink[];
}

@Component({
  selector: 'app-navigation',
  standalone: true,
  imports: [CommonModule, RouterLink],
  template: `
    <aside class="app-sidebar" [class.is-collapsed]="collapsed" [class.is-mobile-open]="mobileOpen" aria-label="Application navigation">
      <a class="sidebar-brand" routerLink="/production" (click)="closeOnMobile()">
        <span class="brand-monogram">S</span>
        <span class="brand-copy">
          <strong>SHAHI</strong>
          <small>Unified Operations</small>
        </span>
      </a>

      <div class="sidebar-rule"></div>
      <nav class="project-navigation" aria-label="Projects">
        <section class="project-group" *ngFor="let project of projects">
          <button
            class="project-toggle"
            type="button"
            [class.project-current]="activeProject === project.id"
            [class.project-expanded]="expandedProject === project.id"
            [attr.aria-expanded]="expandedProject === project.id"
            [attr.aria-label]="project.number + ' ' + project.title"
            [title]="collapsed ? project.title : ''"
            (click)="toggleProject(project)"
          >
            <span class="project-chevron" aria-hidden="true">
              <svg viewBox="0 0 16 16"><path [attr.d]="expandedProject === project.id ? 'm3 6 5 5 5-5' : 'm6 3 5 5-5 5'"></path></svg>
            </span>
            <span class="project-number">{{ project.number }}</span>
            <span class="project-icon" aria-hidden="true">
              <svg *ngIf="project.icon === 'production'" viewBox="0 0 24 24"><path d="M3 21V8l6 3V8l6 3V5l6 3v13H3Z"></path><path d="M7 15h2m3 0h2m3 0h2M7 18h2m3 0h2m3 0h2"></path></svg>
              <svg *ngIf="project.icon === 'cost'" viewBox="0 0 24 24"><path d="M4 19V5m0 14h17M8 15l3-4 3 2 5-7"></path><path d="M16 6h3v3"></path></svg>
              <svg *ngIf="project.icon === 'planning'" viewBox="0 0 24 24"><rect x="5" y="4" width="14" height="17" rx="2"></rect><path d="M9 4.5h6v3H9zM8 12h8m-8 4h5"></path></svg>
            </span>
            <span class="project-copy">
              <span class="project-title">{{ project.title }}</span>
              <span class="project-subtitle">{{ project.subtitle }}</span>
            </span>
          </button>

          <div
            class="project-submenu"
            [class.submenu-expanded]="expandedProject === project.id"
            [attr.aria-hidden]="expandedProject !== project.id"
          >
            <div class="project-submenu-inner">
              <ng-container *ngFor="let link of project.links; let index = index">
                <div
                  class="workflow-section-label"
                  *ngIf="index === 0 || project.links[index - 1].section !== link.section"
                >
                  {{ link.section }}
                </div>
                <a
                  class="workflow-link"
                  [class.workflow-active]="isLinkActive(project, link)"
                  [routerLink]="link.route"
                  [queryParams]="link.queryParams || null"
                  [attr.tabindex]="expandedProject === project.id ? 0 : -1"
                  (click)="closeOnMobile()"
                >
                  <span class="workflow-dot" aria-hidden="true"></span>
                  <span>{{ link.label }}</span>
                </a>
              </ng-container>
            </div>
          </div>
        </section>
      </nav>

      <div class="sidebar-footer">
        <span class="footer-accent"></span>
        <span class="footer-copy">Operations workspace</span>
      </div>

      <button
        class="sidebar-collapse"
        type="button"
        [attr.aria-label]="collapsed ? 'Expand sidebar' : 'Collapse sidebar'"
        [title]="collapsed ? 'Expand sidebar' : 'Collapse sidebar'"
        (click)="collapseToggle.emit()"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path [attr.d]="collapsed ? 'm9 18 6-6-6-6' : 'm15 18-6-6 6-6'"></path></svg>
        <span class="collapse-copy">{{ collapsed ? 'Expand sidebar' : 'Collapse sidebar' }}</span>
      </button>
    </aside>
  `
})
export class AppNavigationComponent implements OnInit, OnDestroy {
  @Input() collapsed = false;
  @Input() mobileOpen = false;
  @Output() collapseToggle = new EventEmitter<void>();
  @Output() mobileClose = new EventEmitter<void>();

  readonly projects: ProjectGroup[] = [
    {
      id: 'production',
      number: '01',
      title: 'Daily DPR',
      subtitle: 'Line Plan Workspace',
      icon: 'production',
      route: '/production',
      links: [
        { section: 'Unit 1', label: 'Plan Generator', route: '/production', queryParams: { unit: 'unit1', mode: 'full' } },
        { section: 'Unit 1', label: 'Daily Actuals Upload', route: '/production', queryParams: { unit: 'unit1', mode: 'daily' } },
        { section: 'Unit 23', label: 'Plan Generator', route: '/production', queryParams: { unit: 'unit23', mode: 'full' } },
        { section: 'Unit 23', label: 'Daily Actuals Upload', route: '/production', queryParams: { unit: 'unit23', mode: 'daily' } }
      ]
    },
    {
      id: 'cost',
      number: '02',
      title: 'India Cost Breakup',
      subtitle: 'Item & Unit Wise Cost Analysis',
      icon: 'cost',
      route: '/cost',
      links: [
        { section: 'India Factory', label: 'Upload Raw', route: '/cost/india/upload' },
        { section: 'India Factory', label: 'Weekly Batches', route: '/cost/india/batches' },
        { section: 'India Factory', label: 'Master Data', route: '/cost/india/master-data' },
        { section: 'Sewtech Factory', label: 'Upload Raw', route: '/cost/sewtech/upload' },
        { section: 'Sewtech Factory', label: 'Weekly Batches', route: '/cost/sewtech/batches' },
        { section: 'Sewtech Factory', label: 'Buyer Mapping & Master', route: '/cost/sewtech/master-data' }
      ]
    },
    {
      id: 'planning',
      number: '03',
      title: 'Planning and Order',
      subtitle: 'Existing Planning & Order Workflows',
      icon: 'planning',
      route: '/planning',
      links: [
        { section: 'Nike', label: 'Snapshot Comparison', route: '/planning/snapshot-comparison' },
        { section: 'Nike', label: 'Category Comparison', route: '/planning/nike-category-comparison' },
        { section: 'Nike', label: 'Order File', route: '/planning/nike-order-file' },
        { section: 'Adidas', label: 'Order File', route: '/planning/adidas-order-file' },
        { section: 'H&M', label: 'Order File', route: '/planning/hm-order-file' },
        { section: 'H&M', label: 'Non-Denim Order File', route: '/planning/hm-order-file-non-denim' },
        { section: 'H&M', label: 'Non-Denim Chart Variation', route: '/planning/hm-non-denim-chart-variation' }
      ]
    }
  ];

  expandedProject: ProjectId | null = null;
  activeProject: ProjectId = 'production';
  private routeSubscription?: Subscription;
  private readonly activeOptions: IsActiveMatchOptions = {
    paths: 'exact',
    queryParams: 'exact',
    fragment: 'ignored',
    matrixParams: 'ignored'
  };

  constructor(private readonly router: Router) {}

  ngOnInit(): void {
    this.syncProjectToRoute();
    this.routeSubscription = this.router.events
      .pipe(filter((event): event is NavigationEnd => event instanceof NavigationEnd))
      .subscribe(() => this.syncProjectToRoute());
  }

  ngOnDestroy(): void {
    this.routeSubscription?.unsubscribe();
  }

  toggleProject(project: ProjectGroup): void {
    if (this.expandedProject === project.id) {
      this.expandedProject = null;
      return;
    }
    this.expandedProject = project.id;
    if (this.activeProject !== project.id) {
      void this.router.navigateByUrl(project.route);
    }
  }

  isLinkActive(project: ProjectGroup, link: WorkflowLink): boolean {
    const tree = this.router.createUrlTree([link.route], { queryParams: link.queryParams });
    if (project.id === 'production') {
      return this.router.isActive(tree, this.activeOptions);
    }
    const currentPath = this.router.url.split('?')[0];
    return currentPath === link.route || currentPath.startsWith(`${link.route}/`);
  }

  closeOnMobile(): void {
    this.mobileClose.emit();
  }

  private syncProjectToRoute(): void {
    const url = this.router.url;
    const project: ProjectId = url.startsWith('/cost')
      ? 'cost'
      : url.startsWith('/planning')
        ? 'planning'
        : 'production';
    this.activeProject = project;
    this.expandedProject = project;
  }
}
