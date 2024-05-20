import { HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';
export interface User {
  _id: number;
  name: string;
  profession: string;
  age: number;
  phoneNumber: number;
  gender:string;
}

@Injectable({
  providedIn: 'root'
})
export class ApiService {
  getUser=new Subject()
  updatedUser = new Subject();
  createdUser=new Subject();
  removeUser=new Subject();
  updateMaterial=new Subject()
  tokenKey="auth"
 constructor(private http:HttpClient) { }
 getMessage(){
    return this.http.get(
        'http://localhost:3001/list');
}
addUser(data:User){
  return this.http.post('http://localhost:3001/addUsers',data)
}
updateUser(data:User){
  return this.http.put(`http://localhost:3001/updateUser`,data)
}
deleteUser(_id:Number){
  return this.http.delete(`http://localhost:3001/deleteUser/${_id}`)
}
addMaterial(data:any){
  return this.http.post("http://[::1]:3000/material-models",data)
  
}
exportMaterial(data:any){
  return this.http.post("http://localhost:4201/exportMaterial",data)

}
updateCabilingMaterial(data:any,id:number){
  return this.http.put(`http://[::1]:3000//material-models/${id}`,data)

}
addingMaterialWithJwt(data:any){
  return this.http.post("http://localhost:4201/sample",data,{
  })
}
storeToken(token:string){
  localStorage.setItem(this.tokenKey,token)

}
getToken():string|null{
  return localStorage.getItem(this.tokenKey)
}
getJwtMaterial(){
  const token=this.getToken()
  const headers=new HttpHeaders().set("Authorization",`Bearer ${token}`)
  return this.http.get('http://localhost:4201/api',{headers:headers})
}

}
