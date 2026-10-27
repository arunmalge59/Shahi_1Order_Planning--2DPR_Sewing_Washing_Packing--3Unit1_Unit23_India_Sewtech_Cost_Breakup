import { Routes } from '@angular/router';
import { ProductionWorkflowComponent } from './modules/production/production-workflow.component';
import { UploadComponent } from './modules/cost/upload/upload.component';
import { BatchListComponent } from './modules/cost/report/batch-list.component';
import { ReportComponent } from './modules/cost/report/report.component';
import { MasterDataComponent } from './modules/cost/buyer-mapping/master-data.component';
import { SewtechUploadComponent } from './modules/cost/sewtech/sewtech-upload.component';
import { SewtechReportComponent } from './modules/cost/sewtech/sewtech-report.component';
import { SewtechFinalComponent } from './modules/cost/sewtech/sewtech-final.component';
import { SewtechBatchesComponent } from './modules/cost/sewtech/sewtech-batches.component';
import { SewtechMasterComponent } from './modules/cost/sewtech/sewtech-master.component';
import { AppComponent as SnapshotComparisonComponent } from './modules/planning/Nike_compare/app.component';
import { NikeCategoryComparisonComponent } from './modules/planning/nike-category-comparison/nike-category-comparison.component';
import { NikeOrderFileComponent } from './modules/planning/nike-order-file.component';
import { AdidasOrderFileComponent } from './modules/planning/adidas-order-file/adidas-order-file.component';
import { HmOrderFileComponent } from './modules/planning/hm-order-file/hm-order-file.component';
import { HmOrderFileNonDenimComponent } from './modules/planning/hm-order-file-non-denim/hm-order-file-non-denim.component';
import { HmNonDenimChartVariationComponent } from './modules/planning/hm-non-denim-chart-variation/hm-non-denim-chart-variation.component';

export const routes: Routes = [
  { path: '', pathMatch: 'full', redirectTo: 'production' },
  { path: 'production', component: ProductionWorkflowComponent },
  { path: 'cost', pathMatch: 'full', redirectTo: 'cost/india/upload' },
  { path: 'cost/india/upload', component: UploadComponent },
  { path: 'cost/india/batches', component: BatchListComponent },
  { path: 'cost/india/batches/:id', component: ReportComponent },
  { path: 'cost/india/master-data', component: MasterDataComponent },
  { path: 'cost/sewtech', pathMatch: 'full', redirectTo: 'cost/sewtech/upload' },
  { path: 'cost/sewtech/upload', component: SewtechUploadComponent },
  { path: 'cost/sewtech/batches', component: SewtechBatchesComponent },
  { path: 'cost/sewtech/report/:id', component: SewtechReportComponent },
  { path: 'cost/sewtech/final/:id', component: SewtechFinalComponent },
  { path: 'cost/sewtech/master-data', component: SewtechMasterComponent },
  { path: 'planning', pathMatch: 'full', redirectTo: 'planning/snapshot-comparison' },
  { path: 'planning/snapshot-comparison', component: SnapshotComparisonComponent },
  { path: 'planning/nike-category-comparison', component: NikeCategoryComparisonComponent },
  { path: 'planning/nike-order-file', component: NikeOrderFileComponent },
  { path: 'planning/adidas-order-file', component: AdidasOrderFileComponent },
  { path: 'planning/hm-order-file', component: HmOrderFileComponent },
  { path: 'planning/hm-order-file-non-denim', component: HmOrderFileNonDenimComponent },
  { path: 'planning/hm-non-denim-chart-variation', component: HmNonDenimChartVariationComponent },
  { path: '**', redirectTo: 'production' }
];
