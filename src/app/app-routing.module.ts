import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { OrderInfoComponent } from './order-info/order-info.component';

const routes: Routes = [
  {path:'Order-Information',component:OrderInfoComponent}
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { 
  
}
