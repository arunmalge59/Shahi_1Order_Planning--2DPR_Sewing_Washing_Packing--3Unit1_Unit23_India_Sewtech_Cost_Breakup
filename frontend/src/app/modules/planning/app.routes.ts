import { Routes } from '@angular/router';
import { AppComponent } from './Nike_compare/app.component';
import { HmOrderFileComponent } from './hm-order-file/hm-order-file.component';
import { HmOrderFileNonDenimComponent } from './hm-order-file-non-denim/hm-order-file-non-denim.component';
import { HmNonDenimChartVariationComponent } from './hm-non-denim-chart-variation/hm-non-denim-chart-variation.component';
import { NikeOrderFileComponent } from './nike-order-file.component';
import { AdidasOrderFileComponent } from './adidas-order-file/adidas-order-file.component';
import { NikeCategoryComparisonComponent } from './nike-category-comparison/nike-category-comparison.component';

export const routes: Routes = [
  { path: '', component: AppComponent },
  { path: 'nike-category-comparison', component: NikeCategoryComparisonComponent },
  { path: 'nike-order-file', component: NikeOrderFileComponent },
  { path: 'adidas-order-file', component: AdidasOrderFileComponent },
  { path: 'hm-order-file', component: HmOrderFileComponent },
  { path: 'hm-order-file-non-denim', component: HmOrderFileNonDenimComponent },
  { path: 'hm-non-denim-chart-variation', component: HmNonDenimChartVariationComponent },
  { path: '**', redirectTo: '' },
];
