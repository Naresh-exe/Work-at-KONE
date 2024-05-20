import { Component } from '@angular/core';
import { ApiService } from '../api.service';
import { ExcelService } from '../excel.service';
import { HttpClient } from '@angular/common/http';
import * as excelJs from 'exceljs'
import * as xmljs from 'xml-js'

@Component({
  selector: 'app-configuration',
  templateUrl: './configuration.component.html',
  styleUrl: './configuration.component.css'
})
export class ConfigurationComponent {
  constructor(private api:ApiService,private excel:ExcelService,private http:HttpClient){}
  ngOnInit(){
   
  }
  
  arr=[]
  cabling={
    machinaryAreaCabiling:[
      {
        elevatorIdentifier:[{A:'',B:'',C:''}],
        materialFor:"Fiber cable from core to building optical patch panel(integrated group) AND Fiber cable from fiber converter to building optical patch panel (zoned group)",
        materialItem:[{value:""},{value:""}],
      },
      {
        elevatorIdentifier:[{ A:'',B:'',C:''}],
        materialFor:"Fiber cable from core 1 to core 2(DCS group cbl core 1 core 2)",
        materialItem:[{value:""},{value:""}],
      },
      {
        elevatorIdentifier:[{ A:'',B:'',C:''}],
        materialFor:"LAN CAT cable from core to GCM core to SC/ER/FCM,and GCM to GCM",
        materialItem:[{value:""},{value:""}],
      },
      {
        elevatorIdentifier:[{ A:'',B:'',C:''}],
        materialFor:"LAN CAT cable from core switch to first access switch in shaft AND LAN CAT cable access switch to access switch in shaft",
        materialItem:[{value:""},{value:""}],
      }
    ]
  }
  col=[
    {
    COLUMN1:{_attributes:{"type":"group_id"}},
    COLUMN2:{_attributes:{"type":"node"}},
    COLUMN3:{_attributes:{"type":"floor"}},
    COLUMN4:{_attributes:{"type":"side"}},
    COLUMN5:{_attributes:{"type":"repeater"}},
    COLUMN6:{_attributes:{"type":"board_id"}},
    COLUMN7:{_attributes:{"type":"board_km"}},
    COLUMN8:{_attributes:{"type":"cable"}},
    COLUMN9:{_attributes:{"type":"cable_length"}}  
  },
  {
    COLUMN1:{_attributes:{"type":"group_id"}},
    COLUMN2:{_attributes:{"type":"node"}},
    COLUMN3:{_attributes:{"type":"floor"}},
    COLUMN4:{_attributes:{"type":"side"}},
    COLUMN5:{_attributes:{"type":"repeater"}},
    COLUMN6:{_attributes:{"type":"board_id"}},
    COLUMN7:{_attributes:{"type":"board_km"}},
    COLUMN8:{_attributes:{"type":"cable"}},
    COLUMN9:{_attributes:{"type":"cable_length"}}  
  },
  {
    COLUMN1:{_attributes:{"type":"group_id"}},
    COLUMN2:{_attributes:{"type":"node"}},
    COLUMN3:{_attributes:{"type":"floor"}},
    COLUMN4:{_attributes:{"type":"side"}},
    COLUMN5:{_attributes:{"type":"repeater"}},
    COLUMN6:{_attributes:{"type":"board_id"}},
    COLUMN7:{_attributes:{"type":"board_km"}},
    COLUMN8:{_attributes:{"type":"cable"}},
    COLUMN9:{_attributes:{"type":"cable_length"}}  
  }
  
  ]
  obj:any={
    ediagram_elevator:{
      LINE1:this.col[0],
      LINE2:this.col[1],
      LINE3:this.col[2]
    },
  }
  
  
  a:any
  condition=false
  arr3:any=[]
  h:any
  info:any
  
   
  save(){
    if(this.condition==false){
    this.api.addingMaterialWithJwt(this.cabling).subscribe((data:any)=>{
      this.cabling=data
      console.log(this.cabling)
      this.h=data.id
      this.condition=true
    })
  }
  else if (this.condition==true){
    this.api.updateCabilingMaterial(this.cabling,this.h).subscribe((data)=>{
      console.log(data)
  
    })
    
  }
  
    }
    export(){
     const workBook=new excelJs.Workbook()
    const workSheet= workBook.addWorksheet('Sheet1')
    const cellA1=workSheet.getCell("A1")
    const cellB1=workSheet.getCell("B1")
    const cellC1=workSheet.getCell("C1")
    cellA1.value="Cable"
    cellB1.value="Cable Material 1"
    cellC1.value="Cable Material 2"
    cellA1.font={bold:true}
    cellB1.font={bold:true}
    cellC1.font={bold:true}
    workSheet.getColumn('A').width=130
    workSheet.getColumn('B').width=30
    workSheet.getColumn('C').width=30
    this.cabling.machinaryAreaCabiling.forEach((obj,index)=>{
    workSheet.getCell(`A${index+2}`).value=obj.materialFor
    this.arr3=[obj.materialItem]
    this.arr3.forEach((obj2:any)=>{
      workSheet.getCell(`B${index+2}`).value=obj2[0].value.toUpperCase()
      workSheet.getCell(`C${index+2}`).value=obj2[1].value.toUpperCase()
      
    })
  })
     workBook.xlsx.writeBuffer().then((data)=>{
       const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url=window.URL.createObjectURL(blob)
       const a=document.createElement('a')
       a.href=url
      a.download='material.xlsx'
      document.body.appendChild(a)
       a.click()
      window.URL.revokeObjectURL(url)
     })
    }
    fileImport(event:any){
      const file:File=event.target.files[0]
      this.excel.readExcel(file)
      this.excel.excelDataTransfer.subscribe((data:any)=>{
      let arr2=[data]
      this.cabling.machinaryAreaCabiling.forEach((obj)=>{
        arr2.forEach((obj2)=>{
          if(obj.materialFor==obj2[0].Cable){
            obj.materialItem[0].value=obj2[0].CableMaterial1
            obj.materialItem[1].value=obj2[0].CableMaterial2
            this.a=obj.materialItem
          }
          else if(obj.materialFor==obj2[1].Cable){
            obj.materialItem[0].value=obj2[1].CableMaterial1
            obj.materialItem[1].value=obj2[1].CableMaterial2
          }
          else if(obj.materialFor==obj2[2].Cable){
            obj.materialItem[0].value=obj2[2].CableMaterial1
            obj.materialItem[1].value=obj2[2].CableMaterial2
          }
          else if(obj.materialFor==obj2[3].Cable){
            obj.materialItem[0].value=obj2[3].CableMaterial1
            obj.materialItem[1].value=obj2[3].CableMaterial2
          } 
        })
      })
     }) 
    
    
  }
  simply(){
    this.api.addingMaterialWithJwt(this.cabling).subscribe((data:any)=>{
       console.log(data)
       this.api.storeToken(data.Token)
       this.info=data.Token
       console.log(this.info)
    })
  
  }
  getMaterial(){
    console.log(this.cabling)
  this.api.getJwtMaterial().subscribe((data)=>{
    console.log(data)
    
  })
  
      }
      down(){
         let xml=xmljs.json2xml(this.obj,{compact:true,spaces:4})
        const blob=new Blob([xml],{type:"application/xml"})
        const url=window.URL.createObjectURL(blob)
        const a=document.createElement('a')
        a.href=url
        a.download="ed.xml"
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        window.URL.revokeObjectURL(url)
  
  
      }
}
