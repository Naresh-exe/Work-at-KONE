import { Component, Inject } from '@angular/core';
import { MAT_DIALOG_DATA } from '@angular/material/dialog';
import { ApiService, User } from '../api.service';
import { MatSnackBar } from '@angular/material/snack-bar';

@Component({
  selector: 'app-delete',
  templateUrl: './delete.component.html',
  styleUrl: './delete.component.css'
})
export class DeleteComponent {
  constructor(@Inject(MAT_DIALOG_DATA) public data: User,private apiService: ApiService,private snackBar:MatSnackBar) { 
  }
  
 
  confirm(){
    this.apiService.deleteUser(this.data._id).subscribe(response=>{
        this.apiService.removeUser.next(response)
     })

    
    this.snackBar.open('Deleted SuccessFully','',{
      duration:3000
    })

}
}
