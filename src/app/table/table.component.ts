import { Component } from '@angular/core';
import { MatTableDataSource } from '@angular/material/table';
import { ApiService, User } from '../api.service';
import { Observable } from 'rxjs';
import { MatDialog } from '@angular/material/dialog';
import { BoxComponent } from '../box/box.component';
import { UpdateComponent } from '../update/update.component';
import { DeleteComponent } from '../delete/delete.component';

@Component({
  selector: 'app-table',
  templateUrl: './table.component.html',
  styleUrl: './table.component.css'
})
export class TableComponent {
  users: Array<User>=[];
  displayedColumns: string[] = ['name', 'work', 'age', 'phoneNumber', 'gender','trah'];
  dataSource: any;


  constructor(private apiService: ApiService, public dial: MatDialog) {
  };
  
  openDialog() {
    this.dial.open(BoxComponent, {
      width: '40%',
    })

  }

  update(data: any) {
    this.dial.open(UpdateComponent, {
      width: '40%',
      data

    })
  }
  Delete(infoData:any){
    this.dial.open(DeleteComponent, {
      width: '20%',
      height:'20%',
      data: infoData

    })
     }
  ngOnInit() {

    this.apiService.getMessage().subscribe((data: any)=>{
      this.users = [...this.users,data];
      this.dataSource = new MatTableDataSource<any>(this.users);
    
    })
     this.apiService.createdUser.subscribe((info:any)=>{
      if (info){
        this.users=[...this.users,info]
        this.dataSource = new MatTableDataSource<any>(this.users);
      }
     })

      this.apiService.updatedUser.subscribe((result: any) => {
       if(result){
        
       this.users.forEach(i=>{
        if(result._id===i._id){
        i._id=result._id
        i.age=result.age
        i.gender=result.gender
        i.name=result.name
        i.phoneNumber=result.phoneNumber
        i.profession=result.profession
        }
       })
       this.dataSource=new MatTableDataSource<any>(this.users)
       }
      

      })
      this.apiService.removeUser.subscribe((res:any)=>{
        let arr:any=[]
        let index=this.users.findIndex(i=>i._id===res._id)
        let rem=this.users.splice(index,1)
        arr=rem
        
        let filter=this.users.filter(data=>{
          return data!=arr
        })
        
        
      this.dataSource=new MatTableDataSource<any>(this.users)
      })    
  }
}
