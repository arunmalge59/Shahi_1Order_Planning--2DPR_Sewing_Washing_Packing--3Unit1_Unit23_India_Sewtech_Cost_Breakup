import { Routes } from '@angular/router';
import { UploadComponent } from './upload/upload.component';
import { BatchListComponent } from './report/batch-list.component';
import { ReportComponent } from './report/report.component';
import { MasterDataComponent } from './buyer-mapping/master-data.component';
import { SewtechUploadComponent } from './sewtech/sewtech-upload.component';
import { SewtechReportComponent } from './sewtech/sewtech-report.component';
import { SewtechFinalComponent } from './sewtech/sewtech-final.component';
import { SewtechBatchesComponent } from './sewtech/sewtech-batches.component';
import { SewtechMasterComponent } from './sewtech/sewtech-master.component';

export const routes: Routes = [
  // Default redirect to India
  { path: '', redirectTo: 'india/upload', pathMatch: 'full' },

  // --- India Routes ---
  { path: 'india', redirectTo: 'india/upload', pathMatch: 'full' },
  { path: 'india/upload', component: UploadComponent },
  { path: 'india/batches', component: BatchListComponent },
  { path: 'india/batches/:id', component: ReportComponent },
  { path: 'india/master-data', component: MasterDataComponent },

  // Legacy India URL redirects & aliases for backward compatibility
  { path: 'upload', component: UploadComponent },
  { path: 'batches', component: BatchListComponent },
  { path: 'batches/:id', component: ReportComponent },
  { path: 'master-data', component: MasterDataComponent },

  // --- Sewtech Routes ---
  { path: 'sewtech', redirectTo: 'sewtech/upload', pathMatch: 'full' },
  { path: 'sewtech/upload', component: SewtechUploadComponent },
  { path: 'sewtech/batches', component: SewtechBatchesComponent },
  { path: 'sewtech/report/:id', component: SewtechReportComponent },
  { path: 'sewtech/final/:id', component: SewtechFinalComponent },
  { path: 'sewtech/master-data', component: SewtechMasterComponent },

  // Wildcard fallback
  { path: '**', redirectTo: 'india/upload' },
];
