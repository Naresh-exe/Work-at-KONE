import { Component } from '@angular/core';
import { FormControl, FormGroup, Validators } from '@angular/forms';
import { MatSnackBar } from '@angular/material/snack-bar';
import { ApiService } from '../api.service';

@Component({
  selector: 'app-box',
  templateUrl: './box.component.html',
  styleUrl: './box.component.css'
})
export class BoxComponent {
  constructor( private snackBar:MatSnackBar,private apiService:ApiService){}
  
  index=new FormGroup({
    id:new FormControl(''),
   name:new FormControl('',[Validators.required,Validators.minLength(2)]),
   profession:new FormControl('',Validators.required),
   age:new FormControl('',[Validators.maxLength(2),Validators.pattern('[- +()0-9]+'),Validators.required]),
   phoneNumber:new FormControl('',[Validators.pattern('[- +()0-9]+'),Validators.required]) ,
   gender:new FormControl('',Validators.required)

  })    
info:any
Submit(){
this.info=this.index.value;
this.apiService.addUser(this.info).subscribe(data=>{

this.apiService.createdUser.next(data)

this.snackBar.open('Created Successfully','',{
duration:3000
})
})

}

}
