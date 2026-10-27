import { Component, OnDestroy, OnInit } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { Subscription } from 'rxjs';
import { AppComponent as ProductionComponent } from './app.component';

@Component({
  selector: 'app-production-workflow',
  standalone: true,
  imports: [ProductionComponent],
  template: `
    <section class="unified-production-page">
      <app-production-workspace
        [externalNavigation]="true"
        [workflowUnit]="unit"
        [workflowMode]="mode"
        (workflowChange)="navigateWorkflow($event)"
      ></app-production-workspace>
    </section>
  `
})
export class ProductionWorkflowComponent implements OnInit, OnDestroy {
  private routeSubscription?: Subscription;
  unit: 'unit1' | 'unit23' = 'unit1';
  mode: 'full' | 'daily' = 'full';

  constructor(private readonly route: ActivatedRoute, private readonly router: Router) {
    const params = new URLSearchParams(window.location.search);
    this.unit = params.get('unit') === 'unit23' ? 'unit23' : 'unit1';
    this.mode = params.get('mode') === 'daily' ? 'daily' : 'full';
  }

  ngOnInit(): void {
    this.routeSubscription = this.route.queryParamMap.subscribe((params) => {
      this.unit = params.get('unit') === 'unit23' ? 'unit23' : 'unit1';
      this.mode = params.get('mode') === 'daily' ? 'daily' : 'full';
    });
  }

  ngOnDestroy(): void {
    this.routeSubscription?.unsubscribe();
  }

  navigateWorkflow(workflow: { unit: 'unit1' | 'unit23'; mode: 'full' | 'daily' }): void {
    if (workflow.unit === this.unit && workflow.mode === this.mode) return;
    this.unit = workflow.unit;
    this.mode = workflow.mode;
    void this.router.navigate([], {
      relativeTo: this.route,
      queryParams: { unit: workflow.unit, mode: workflow.mode },
      queryParamsHandling: 'merge',
      replaceUrl: true
    });
  }
}
