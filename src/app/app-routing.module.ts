import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { OrderInfoComponent } from './order-info/order-info.component';
import { ConfigurationComponent } from './configuration/configuration.component';
import { TableComponent } from './table/table.component';

const routes: Routes = [
  {path:'Order-Information',component:OrderInfoComponent},{path:'configuration',component:ConfigurationComponent},{path:'CRUD',component:TableComponent}
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { 
  
}
