import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';
import * as XLSX from 'xlsx'

@Injectable({
  providedIn: 'root'
})
export class ExcelService {
  constructor(private http:HttpClient) { }

  excelDataTransfer=new Subject()
  postFile(data:any){
  return this.http.post('http://localhost:4100/hello',data)

  }
 readExcel(file:File){
   const reader=new FileReader()
   reader.onload=(dataInExcel:any)=>{
     const data=dataInExcel.target.result
     const workBook=XLSX.read(data,{type:'binary'})
     const sheetName=workBook.SheetNames[0]
     const workSheet=workBook.Sheets[sheetName]
     const excelData=XLSX.utils.sheet_to_json(workSheet,{raw:true})
     this.excelDataTransfer.next(excelData)
   

   }
   reader.readAsBinaryString(file)
 }
 uploadFile(file:File){
  const formData=new FormData()
  formData.append('helod',file)
 }
 xmlfile(file:File){
   const reader=new FileReader()
   reader.onload=(dataInExcel:any)=>{
     const data=dataInExcel.target.result
     const workBook=XLSX.read(data,{type:'binary'})
     const sheetName=workBook.SheetNames[0]
     const workSheet=workBook.Sheets[sheetName]
     const xmlData:any=XLSX.utils.sheet_to_json(workSheet)
     
     
   }
   reader.readAsBinaryString(file)
 }
}
