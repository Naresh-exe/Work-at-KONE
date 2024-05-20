import { Component } from '@angular/core';
import * as XLSX from 'xlsx'
import { ExcelService } from '../excel.service';

@Component({
  selector: 'app-order-info',
  templateUrl: './order-info.component.html',
  styleUrl: './order-info.component.css'
})
export class OrderInfoComponent {
  constructor(private ex:ExcelService){}
  arrofD880:any=[]
  arrOfFlowChartD880:any=[]
  arrOfD1020:any=[]
  arrofFlowchartD1020:any=[]
  visualGuidance:any=[]
  walkingTimes:any=[]
  distance={
    distanceFromCore1ToCore2:'',
    distanceFromCore1ToSiteController:"",
    distanceFromCore2ToSiteController:"",
    distanceFromCore1ToEdgeRouter1:"",
    distanceFromCore2ToEdgeRouter1:'',
    distanceFromCore2ToEdgeRouter2:"",
    distanceFromCore1To1stGCM:"",
    distanceFromCore2ToLastGCM:"",
    distanceFromEdgeRouter1ToBuildingLAN:"",
    distanceFromCore2ToFiberConverter2:""
  }
  cableDistance={
    cableFromGCMToController:{
      G:'',H:'',I:'',J:"",K:'',L:''
    },
    cableFromGCMToServedFloorClosestToMachineRoom:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    cableFromGCMToGCM:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    CoreSwitchElevator:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    LANRiserElevators:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    cableFromCoreSwitchToFirstAccessSwitchInShaft:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    distanceToBuildingOpticalpatchPanelOfIntegratedGroup:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    siteControllerLocation:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    cableFromAMSMToGCMAndIsolationSwitchToGCM:
      {G:'', H:"", I:'', J:'', K:"", L:""},
    cableFromAMSMToIsolationSwitch:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    AMSMToCoreSwitchandIsolationSwitchToCoreSwitch:
      { G:'', H:"", I:'', J:'', K:"", L:""},
    extensionCableqty:
      { G:'', H:"", I:'', J:'', K:"", L:""}
    
  }
  walkingTimeObj=[
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
  {
    elevatorIdentifier:'',
    dopIdentifier:'',
    distance:""
  },
 

]
  visualGuidanceObj=[
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
    {
      elevatorIdentifier:"",
      dopIdentifier:'',
      distance:""
    },
  ]
  d880={
    orderInformation:{
        salesPerson:'',
        flUnit:'',
        Date:'',
        phone:'',
        email:'',
        networkNumber:'',
        elevatorNumber:'',
        koneEquipmentNumber:'',
        koneOpportunityNumber:'',
        ESUReferenceNumber:'',
        projectName:'',
        reciever:'',
        streetAddress:'',
        zipCode:'',
        city:'',
        country:'',
        revDate:'',
        order:''
    },
    Elevators:[
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:0
        },
        dops:[{
          dopIdentifier:''
        }]
      },
       {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:1
        },
        dops:[{
          dopIdentifier:''
        }]
      },  {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:2
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:3
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:4
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:5
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:6
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:7
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:8
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:9
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:10
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:11
        },
        dops:[{
          dopIdentifier:''
        }]
      },
      {
        elevatorIdentifier:'',
        position:{
          x:0,
          y:12
        },
        dops:[{
          dopIdentifier:''
        }]
      }
    ],
    floors:[
      {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
 
  ],
  }
  d1020={
    orderInformation:{
      flSalesPerson:'',
      flUnit:'',
      date:'',
      phone:'',
      email:'',
      bussinesstype:'',
      projectName:'',
      siteCountry:'',
      siteStreetAddress:'',
      supplyLine:'',
      koneOppurtunityNumber:'',
      kpBinderNumber:'',
      revDate:''
    },
      elevators: [
        {
            elevatorIdentifier: '',
            position: {
                x: 0,
                y: 0
            },
            dops: [{
                dopIdentifier: '',
            }]
        },
        {
            elevatorIdentifier: '',
            dops: [{
              dopIdentifier: '',
          }],
            position: {
                x: 0,
                y: 1
            }
        },
        {
            elevatorIdentifier: '',
            dops: [{
              dopIdentifier: '',
          }],
            position: {
                x: 1,
                y: 0
            }
        }
    ],
  floors:[
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
    {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
     {
      floorNo:'',
      floorMaking:'',
      serveASide:'',
      serveBSide:''
    },
  
  ]
  }

  import(event:any){
    const fileSelector=event.target.files[0]
    const fileReader=new FileReader()
    fileReader.readAsBinaryString(fileSelector)
    fileReader.onload=(event:any)=>{
      const binaryData=event.target.result
      let workBook=XLSX.read(binaryData,{type:'binary'})
      const flowChartsheetName=workBook.SheetNames[15]
      workBook.SheetNames.forEach((sheet)=>{
        const sheets=workBook.Sheets[sheet]
        const allData=XLSX.utils.sheet_to_json(sheets)
        allData.forEach((value:any)=>{
          let ob=Object.values(value)
          this.arrofD880.push(ob)
        })

      })
      const sheet=workBook.Sheets[flowChartsheetName]
      const flowChartdata= XLSX.utils.sheet_to_json(sheet)
      flowChartdata.forEach((value:any)=>{
        let obj=Object.values(value)
        this.arrOfFlowChartD880.push(obj)
      })
      if(fileSelector.name==="D880.xlsm"){
      this.d880.orderInformation.salesPerson=this.arrofD880[4][2]
      this.d880.orderInformation.flUnit=this.arrofD880[5][1]
      this.d880.orderInformation.Date=this.arrofD880[6][1]
      this.d880.orderInformation.phone=this.arrofD880[7][2]
      this.d880.orderInformation.email=this.arrofD880[9][1]
      this.d880.orderInformation.networkNumber=this.arrofD880[10][2]
      this.d880.orderInformation.elevatorNumber=this.arrofD880[11][1]
      this.d880.orderInformation.koneEquipmentNumber=this.arrofD880[12][1]
      this.d880.orderInformation.koneOpportunityNumber=this.arrofD880[14][1]
      this.d880.orderInformation.ESUReferenceNumber=this.arrofD880[15][1]
      this.d880.orderInformation.projectName=this.arrofD880[16][2]
      this.d880.orderInformation.reciever=this.arrofD880[17][1]
      this.d880.orderInformation.streetAddress=this.arrofD880[18][1]
      this.d880.orderInformation.zipCode=this.arrofD880[19][1]
      this.d880.orderInformation.city=this.arrofD880[20][1]
      this.d880.orderInformation.country=this.arrofD880[21][1]
      this.d880.orderInformation.revDate=this.arrofD880[22][2]
      this.d880.floors.forEach((data,index)=>{
        if(index==0){
        data.floorNo=this.arrOfFlowChartD880[169][0]
        data.floorMaking=this.arrOfFlowChartD880[169][2]
        data.serveASide=this.arrOfFlowChartD880[169][3]
        data.serveBSide=this.arrOfFlowChartD880[169][4]
        }
        else if(index==1){
          data.floorNo=this.arrOfFlowChartD880[167][0]
          data.floorMaking=this.arrOfFlowChartD880[167][2]
          data.serveASide=this.arrOfFlowChartD880[167][3]
          data.serveBSide=this.arrOfFlowChartD880[167][4]

        }
        else if(index==2){
          data.floorNo=this.arrOfFlowChartD880[165][0]
          data.floorMaking=this.arrOfFlowChartD880[165][2]
          data.serveASide=this.arrOfFlowChartD880[165][3]
          data.serveBSide=this.arrOfFlowChartD880[165][4]
        }
        else if(index==3){
          data.floorNo=this.arrOfFlowChartD880[163][0]
          data.floorMaking=this.arrOfFlowChartD880[163][2]
          data.serveASide=this.arrOfFlowChartD880[163][3]
          data.serveBSide=this.arrOfFlowChartD880[163][4]
        }
         else if(index==4){
          data.floorNo=this.arrOfFlowChartD880[161][0]
          data.floorMaking=this.arrOfFlowChartD880[161][2]
          data.serveASide=this.arrOfFlowChartD880[161][3]
          data.serveBSide=this.arrOfFlowChartD880[161][4]
        }
         else if(index==5){
          data.floorNo=this.arrOfFlowChartD880[159][0]
          data.floorMaking=this.arrOfFlowChartD880[159][2]
          data.serveASide=this.arrOfFlowChartD880[159][3]
          data.serveBSide=this.arrOfFlowChartD880[159][4]
        }
          else if(index==6){
          data.floorNo=this.arrOfFlowChartD880[157][0]
          data.floorMaking=this.arrOfFlowChartD880[157][2]
          data.serveASide=this.arrOfFlowChartD880[157][3]
          data.serveBSide=this.arrOfFlowChartD880[157][4]
        }
          else if(index==7){
          data.floorNo=this.arrOfFlowChartD880[155][0]
          data.floorMaking=this.arrOfFlowChartD880[155][2]
          data.serveASide=this.arrOfFlowChartD880[155][3]
          data.serveBSide=this.arrOfFlowChartD880[155][4]
        }
          else if(index==8){
          data.floorNo=this.arrOfFlowChartD880[153][0]
          data.floorMaking=this.arrOfFlowChartD880[153][2]
          data.serveASide=this.arrOfFlowChartD880[153][3]
          data.serveBSide=this.arrOfFlowChartD880[153][4]
        }
          else if(index==9){
          data.floorNo=this.arrOfFlowChartD880[151][0]
          data.floorMaking=this.arrOfFlowChartD880[151][2]
          data.serveASide=this.arrOfFlowChartD880[151][3]
          data.serveBSide=this.arrOfFlowChartD880[151][4]
        }
          else if(index==10){
          data.floorNo=this.arrOfFlowChartD880[149][0]
          data.floorMaking=this.arrOfFlowChartD880[149][2]
          data.serveASide=this.arrOfFlowChartD880[149][3]
          data.serveBSide=this.arrOfFlowChartD880[149][4]
        }
         else if(index==11){
          data.floorNo=this.arrOfFlowChartD880[148][0]
          data.floorMaking=this.arrOfFlowChartD880[148][2]
          data.serveASide=this.arrOfFlowChartD880[148][3]
          data.serveBSide=this.arrOfFlowChartD880[148][4]
        }
         else if(index==12){
          data.floorNo=this.arrOfFlowChartD880[146][0]
          data.floorMaking=this.arrOfFlowChartD880[146][2]
          data.serveASide=this.arrOfFlowChartD880[146][3]
          data.serveBSide=this.arrOfFlowChartD880[146][4]
        }
         else if(index==13){
          data.floorNo=this.arrOfFlowChartD880[144][0]
          data.floorMaking=this.arrOfFlowChartD880[144][2]
          data.serveASide=this.arrOfFlowChartD880[144][3]
          data.serveBSide=this.arrOfFlowChartD880[144][4]
        }
        
      })
      this.d880.Elevators.forEach((value,index)=>{
        if(index==0){
          value.elevatorIdentifier=this.arrOfFlowChartD880[18][1]
          value.dops.forEach((data)=>{
            data.dopIdentifier=this.arrOfFlowChartD880[22][2]
           })
        }
        else if(index==1){
          value.elevatorIdentifier=this.arrOfFlowChartD880[18][3]
          value.dops.forEach((data)=>{
            data.dopIdentifier=this.arrOfFlowChartD880[22][4]
          })
        }
        else if(index==2){
          value.elevatorIdentifier=this.arrOfFlowChartD880[18][5]
          value.dops.forEach((data)=>{
            data.dopIdentifier=this.arrOfFlowChartD880[22][4]
          })
        }
        else if(index==3){
        }
        })
        console.log(this.arrOfFlowChartD880)
      
      }
      else if(fileSelector.name==="D1020.xlsm"){
       
        //d1020
        const orderInfoOfd1020SheetName=workBook.SheetNames[1]
        const flowChartD1020SheetName=workBook.SheetNames[8]
        const sheets=workBook.Sheets[orderInfoOfd1020SheetName]
        const sheets1=workBook.Sheets[flowChartD1020SheetName]
        const orderInfoData=XLSX.utils.sheet_to_json(sheets)
        const flowChartdataOfD1020=XLSX.utils.sheet_to_json(sheets1)
        orderInfoData.forEach((value:any)=>{
          let obj=Object.values(value)
          this.arrOfD1020.push(obj)
        })
        flowChartdataOfD1020.forEach((value:any)=>{
          let obj1=Object.values(value)
          this.arrofFlowchartD1020.push(obj1)
        })
        
        this.d1020.orderInformation.flSalesPerson=this.arrOfD1020[4][1] || null
        this.d1020.orderInformation.flUnit=this.arrOfD1020[5][1] || null
        this.d1020.orderInformation.date=''
        this.d1020.orderInformation.phone=''
        this.d1020.orderInformation.email=this.arrOfD1020[8][1] || null
        this.d1020.orderInformation.bussinesstype=this.arrOfD1020[11][1] || null
        this.d1020.orderInformation.projectName=this.arrOfD1020[13][1] || null
        this.d1020.orderInformation.siteCountry=this.arrOfD1020[14][1] || null
        this.d1020.orderInformation.siteStreetAddress=this.arrOfD1020[15][1] || null
        this.d1020.orderInformation.supplyLine=this.arrOfD1020[16][1] || null
        this.d1020.orderInformation.koneOppurtunityNumber=this.arrOfD1020[17][1] || null
        this.d1020.orderInformation.kpBinderNumber=this.arrOfD1020[20][1] || null
        this.d1020.orderInformation.revDate=this.arrOfD1020[25][1] || null
        this.d1020.elevators.forEach((value,index)=>{
          if(index==0){
            value.elevatorIdentifier=this.arrofFlowchartD1020[13][4] || null
            value.dops.forEach((data)=>{
              data.dopIdentifier=this.arrofFlowchartD1020[18][1] || null
            })
          }
          else if(index==1){
            value.elevatorIdentifier=this.arrofFlowchartD1020[13][7] || null
            value.dops.forEach((data)=>{
              data.dopIdentifier=this.arrofFlowchartD1020[18][2] || null
            })
          }
          
        })
        this.d1020.floors.forEach((data,index)=>{
          if(index==0){
            data.floorNo=this.arrofFlowchartD1020[206][0] || null
            data.floorMaking=this.arrofFlowchartD1020[206][1] || null
            data.serveASide=this.arrofFlowchartD1020[206][3] || null
          }
         else if(index==1){
            data.floorNo=this.arrofFlowchartD1020[205][0] || null
            data.floorMaking=this.arrofFlowchartD1020[205][1] || null
            data.serveASide=this.arrofFlowchartD1020[205][3] || null
          }
         else if(index==2){
            data.floorNo=this.arrofFlowchartD1020[204][0] || null
            data.floorMaking=this.arrofFlowchartD1020[204][1]|| null
            data.serveASide=this.arrofFlowchartD1020[204][3] || null
          }
         else if(index==3){
            data.floorNo=this.arrofFlowchartD1020[203][0] || null
            data.floorMaking=this.arrofFlowchartD1020[203][1] || null
            data.serveASide=this.arrofFlowchartD1020[203][3] || null
          }
        else  if(index==4){
            data.floorNo=this.arrofFlowchartD1020[202][0] || null
            data.floorMaking=this.arrofFlowchartD1020[202][1] || null
            data.serveASide=this.arrofFlowchartD1020[202][3] || null
          }
         else if(index==5){
            data.floorNo=this.arrofFlowchartD1020[201][0] || null
            data.floorMaking=this.arrofFlowchartD1020[201][1] || null
            data.serveASide=this.arrofFlowchartD1020[201][3] || null
          }
         else if(index==6){
            data.floorNo=this.arrofFlowchartD1020[200][0] || null
            data.floorMaking=this.arrofFlowchartD1020[200][1] || null
            data.serveASide=this.arrofFlowchartD1020[200][3] || null
          }
          else if(index==7){
            data.floorNo=this.arrofFlowchartD1020[199][0] || null
            data.floorMaking=this.arrofFlowchartD1020[199][1] || null
            data.serveASide=this.arrofFlowchartD1020[199][3] || null
          }
         else if(index==8){
            data.floorNo=this.arrofFlowchartD1020[198][0] || null
            data.floorMaking=this.arrofFlowchartD1020[198][1] || null
            data.serveASide=this.arrofFlowchartD1020[198][3] || null
          }
         else if(index==9){
            data.floorNo=this.arrofFlowchartD1020[197][0]|| null
            data.floorMaking=this.arrofFlowchartD1020[197][1]|| null
            data.serveASide=this.arrofFlowchartD1020[197][3] || null
          }
          else if(index==10){
            data.floorNo=this.arrofFlowchartD1020[196][0] || null
            data.floorMaking=this.arrofFlowchartD1020[196][1] || null
            data.serveASide=this.arrofFlowchartD1020[196][3] || null
          }
          else if(index==11){
            data.floorNo=this.arrofFlowchartD1020[195][0] || null
            data.floorMaking=this.arrofFlowchartD1020[195][1] || null
            data.serveASide=this.arrofFlowchartD1020[195][3] || null
          }
          else if(index==12){
            data.floorNo=this.arrofFlowchartD1020[194][0] || null
            data.floorMaking=this.arrofFlowchartD1020[194][1] || null
            data.serveASide=this.arrofFlowchartD1020[194][3] || null
          }
         else if(index==13){
            data.floorNo=this.arrofFlowchartD1020[193][0] || null
            data.floorMaking=this.arrofFlowchartD1020[193][1] || null
            data.serveASide=this.arrofFlowchartD1020[193][3] || null
          }
         else if(index==14){
            data.floorNo=this.arrofFlowchartD1020[192][0] || null
            data.floorMaking=this.arrofFlowchartD1020[192][1] || null
            data.serveASide=this.arrofFlowchartD1020[192][3] || null
          }
           else if(index==15){
            data.floorNo=this.arrofFlowchartD1020[191][0] || null
            data.floorMaking=this.arrofFlowchartD1020[191][1] || null
            data.serveASide=this.arrofFlowchartD1020[191][3] || null
          }
           else if(index==16){
            data.floorNo=this.arrofFlowchartD1020[190][0] || null
            data.floorMaking=this.arrofFlowchartD1020[190][1] || null
            data.serveASide=this.arrofFlowchartD1020[190][3]|| null
          }
           else if(index==17){
            data.floorNo=this.arrofFlowchartD1020[189][0] || null
            data.floorMaking=this.arrofFlowchartD1020[189][1] || null
            data.serveASide=this.arrofFlowchartD1020[189][3] || null
          }
           else if(index==18){
            data.floorNo=this.arrofFlowchartD1020[188][0] || null
            data.floorMaking=this.arrofFlowchartD1020[188][1] || null
            data.serveASide=this.arrofFlowchartD1020[188][3] || null
          }
           else if(index==19){
            data.floorNo=this.arrofFlowchartD1020[187][0] || null
            data.floorMaking=this.arrofFlowchartD1020[187][1] || null
            data.serveASide=this.arrofFlowchartD1020[187][3] || null
          }
           else if(index==20){
            data.floorNo=this.arrofFlowchartD1020[186][0] || null
            data.floorMaking=this.arrofFlowchartD1020[186][1] || null
            data.serveASide=this.arrofFlowchartD1020[186][3] || null
          }
           else if(index==21){
            data.floorNo=this.arrofFlowchartD1020[185][0] || null
            data.floorMaking=this.arrofFlowchartD1020[185][1] || null
            data.serveASide=this.arrofFlowchartD1020[185][3] || null
          }
           else if(index==22){
            data.floorNo=this.arrofFlowchartD1020[184][0] || null
            data.floorMaking=this.arrofFlowchartD1020[184][1] || null
            data.serveASide=this.arrofFlowchartD1020[184][3] || null
          }
           else if(index==23){
            data.floorNo=this.arrofFlowchartD1020[183][0] || null
            data.floorMaking=this.arrofFlowchartD1020[183][1] || null
            data.serveASide=this.arrofFlowchartD1020[183][3] || null
          }
           else if(index==24){
            data.floorNo=this.arrofFlowchartD1020[182][0] || null
            data.floorMaking=this.arrofFlowchartD1020[182][1] || null
            data.serveASide=this.arrofFlowchartD1020[182][3] || null
          }
           else if(index==25){
            data.floorNo=this.arrofFlowchartD1020[181][0] || null
            data.floorMaking=this.arrofFlowchartD1020[181][1] || null
            data.serveASide=this.arrofFlowchartD1020[181][3] || null
          }
           else if(index==26){
            data.floorNo=this.arrofFlowchartD1020[180][0] || null
            data.floorMaking=this.arrofFlowchartD1020[180][1] || null
            data.serveASide=this.arrofFlowchartD1020[180][3] || null
          }
           else if(index==27){
            data.floorNo=this.arrofFlowchartD1020[179][0] || null
            data.floorMaking=this.arrofFlowchartD1020[179][1] || null
            data.serveASide=this.arrofFlowchartD1020[179][3] || null
          }
           else if(index==28){
            data.floorNo=this.arrofFlowchartD1020[178][0] || null
            data.floorMaking=this.arrofFlowchartD1020[178][1] || null
            data.serveASide=this.arrofFlowchartD1020[178][3] || null
          }
           else if(index==29){
            data.floorNo=this.arrofFlowchartD1020[177][0] || null
            data.floorMaking=this.arrofFlowchartD1020[177][1] || null
            data.serveASide=this.arrofFlowchartD1020[177][3] || null
          }
           else if(index==30){
            data.floorNo=this.arrofFlowchartD1020[176][0] || null
            data.floorMaking=this.arrofFlowchartD1020[176][1] || null
            data.serveASide=this.arrofFlowchartD1020[176][3] || null
          }
           else if(index==31){
            data.floorNo=this.arrofFlowchartD1020[175][0] || null
            data.floorMaking=this.arrofFlowchartD1020[175][1] || null
            data.serveASide=this.arrofFlowchartD1020[175][3] || null
          }
           else if(index==32){
            data.floorNo=this.arrofFlowchartD1020[174][0] || null
            data.floorMaking=this.arrofFlowchartD1020[174][1] || null
            data.serveASide=this.arrofFlowchartD1020[174][3] || null
          }
           else if(index==33){
            data.floorNo=this.arrofFlowchartD1020[173][0] || null
            data.floorMaking=this.arrofFlowchartD1020[173][1] || null
            data.serveASide=this.arrofFlowchartD1020[173][3] || null
          }
           else if(index==34){
            data.floorNo=this.arrofFlowchartD1020[172][0] || null
            data.floorMaking=this.arrofFlowchartD1020[172][1] || null
            data.serveASide=this.arrofFlowchartD1020[172][3] || null
          }
           else if(index==35){
            data.floorNo=this.arrofFlowchartD1020[171][0] || null
            data.floorMaking=this.arrofFlowchartD1020[171][1] || null
            data.serveASide=this.arrofFlowchartD1020[171][3] || null
          }
           else if(index==36){
            data.floorNo=this.arrofFlowchartD1020[170][0] || null
            data.floorMaking=this.arrofFlowchartD1020[170][1] || null
            data.serveASide=this.arrofFlowchartD1020[170][3] || null
          }
        }) 
      this.walkingTimeObj.forEach((value,index)=>{
        if(index==0){
          value.dopIdentifier=this.arrofFlowchartD1020[51][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][0] || null
          value.distance=this.arrofFlowchartD1020[51][1] || null
        }
        else if(index==1){ 
          value.dopIdentifier=this.arrofFlowchartD1020[51][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][1] || null
          value.distance=this.arrofFlowchartD1020[51][2] || null
        }
        else if(index==2){
          value.dopIdentifier=this.arrofFlowchartD1020[51][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][2] || null
          value.distance=this.arrofFlowchartD1020[51][3] || null
        }
        else if(index==3){
          value.dopIdentifier=this.arrofFlowchartD1020[51][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][3] || null
          value.distance=this.arrofFlowchartD1020[51][4] || null
        }
        else if(index==4){
          value.dopIdentifier=this.arrofFlowchartD1020[51][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][4] || null
          value.distance=this.arrofFlowchartD1020[51][5] || null
        }
        else if(index==5){
          value.dopIdentifier=this.arrofFlowchartD1020[51][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][5] || null
          value.distance=this.arrofFlowchartD1020[51][6] || null
        }
        else if(index==6){
          value.dopIdentifier=this.arrofFlowchartD1020[52][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][0] || null
          value.distance=this.arrofFlowchartD1020[52][1] || null
        }
        else if(index==7){
          value.dopIdentifier=this.arrofFlowchartD1020[52][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][1] || null
          value.distance=this.arrofFlowchartD1020[52][2] || null
        }
        else if(index==8){
          value.dopIdentifier=this.arrofFlowchartD1020[52][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][2] || null
          value.distance=this.arrofFlowchartD1020[52][3] || null
        }
        else if(index==9){
          value.dopIdentifier=this.arrofFlowchartD1020[52][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][3] || null
          value.distance=this.arrofFlowchartD1020[52][4] || null
        }
        else if(index==10){
          value.dopIdentifier=this.arrofFlowchartD1020[52][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][4] || null
          value.distance=this.arrofFlowchartD1020[52][5] || null
        }
        else if(index==11){
          value.dopIdentifier=this.arrofFlowchartD1020[52][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][5] || null
          value.distance=this.arrofFlowchartD1020[52][6] || null
        }
        else if(index==12){
          value.dopIdentifier=this.arrofFlowchartD1020[53][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][0] || null
          value.distance=this.arrofFlowchartD1020[53][1] || null
        }
        else if(index==13){
          value.dopIdentifier=this.arrofFlowchartD1020[53][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][1] || null
          value.distance=this.arrofFlowchartD1020[53][2] || null
        }
        else if(index==14){
          value.dopIdentifier=this.arrofFlowchartD1020[53][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][2] || null
          value.distance=this.arrofFlowchartD1020[53][3] || null
        }
        else if(index==15){
          value.dopIdentifier=this.arrofFlowchartD1020[53][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][3] || null
          value.distance=this.arrofFlowchartD1020[53][4] || null
        }
        else if(index==16){ 
          value.dopIdentifier=this.arrofFlowchartD1020[53][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][4] || null
          value.distance=this.arrofFlowchartD1020[53][5] || null
        }
        else if(index==17){
          value.dopIdentifier=this.arrofFlowchartD1020[53][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][5] || null
          value.distance=this.arrofFlowchartD1020[53][6] || null
        }
        else if(index==18){
          value.dopIdentifier=this.arrofFlowchartD1020[54][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][0] || null
          value.distance=this.arrofFlowchartD1020[54][1] || null
        }
        else if(index==19){
          value.dopIdentifier=this.arrofFlowchartD1020[54][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][1] || null
          value.distance=this.arrofFlowchartD1020[54][2] || null
        }
        else if(index==20){
          value.dopIdentifier=this.arrofFlowchartD1020[54][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][2] || null
          value.distance=this.arrofFlowchartD1020[54][3] || null
        }
        else if(index==21){
          value.dopIdentifier=this.arrofFlowchartD1020[54][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][3] || null
          value.distance=this.arrofFlowchartD1020[54][4] || null
        }
        else if(index==22){
          value.dopIdentifier=this.arrofFlowchartD1020[54][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][4] || null
          value.distance=this.arrofFlowchartD1020[54][5] || null
        }
        else if(index==23){
          value.dopIdentifier=this.arrofFlowchartD1020[54][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][5] || null
          value.distance=this.arrofFlowchartD1020[54][6] || null
        }
        else if(index==24){
          value.dopIdentifier=this.arrofFlowchartD1020[55][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][0] || null
          value.distance=this.arrofFlowchartD1020[55][1] || null
        }
        else if(index==25){
          value.dopIdentifier=this.arrofFlowchartD1020[55][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][1] || null
          value.distance=this.arrofFlowchartD1020[55][2] || null
        }
        else if(index==26){
          value.dopIdentifier=this.arrofFlowchartD1020[55][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][2] || null
          value.distance=this.arrofFlowchartD1020[55][3] || null
        }
        else if(index==27){
          value.dopIdentifier=this.arrofFlowchartD1020[55][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][3] || null
          value.distance=this.arrofFlowchartD1020[55][4] || null
        }
        else if(index==28){
          value.dopIdentifier=this.arrofFlowchartD1020[55][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][4] || null
          value.distance=this.arrofFlowchartD1020[55][5] || null
        }
        else if(index==29){
          value.dopIdentifier=this.arrofFlowchartD1020[55][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][5] || null
          value.distance=this.arrofFlowchartD1020[55][6] || null
        }
        else if(index==30){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][0] || null
          value.distance=this.arrofFlowchartD1020[56][1] || null
        }
        else if(index==31){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][1] || null
          value.distance=this.arrofFlowchartD1020[56][2] || null
        }
        else if(index==32){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][2] || null
          value.distance=this.arrofFlowchartD1020[56][3] || null
        }
        else if(index==33){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][3] || null
          value.distance=this.arrofFlowchartD1020[56][4] || null
        }
        else if(index==34){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][4] || null
          value.distance=this.arrofFlowchartD1020[56][5] || null
        }
        else if(index==35){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0] || null
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][5] || null
          value.distance=this.arrofFlowchartD1020[56][6] || null
        }
        else if(index==36){
          value.dopIdentifier=this.arrofFlowchartD1020[56][0]
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][6]
          value.distance=this.arrofFlowchartD1020[56][7]
        }
      })
      this.visualGuidanceObj.forEach((value,index)=>{
        if(index==0){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][12]
          value.dopIdentifier=this.arrofFlowchartD1020[51][13]
          value.distance=this.arrofFlowchartD1020[51][14]
        }
        else if(index==1){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][13]
          value.dopIdentifier=this.arrofFlowchartD1020[51][13]
          value.distance=this.arrofFlowchartD1020[51][15]
        }
        else if(index==2){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][14]
          value.dopIdentifier=this.arrofFlowchartD1020[51][13]
          value.distance=this.arrofFlowchartD1020[51][16]
        }
        else if(index==3){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][15]
          value.dopIdentifier=this.arrofFlowchartD1020[51][13]
          value.distance=this.arrofFlowchartD1020[51][17]
        }
        else if(index==4){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][16]
          value.dopIdentifier=this.arrofFlowchartD1020[51][13]
          value.distance=this.arrofFlowchartD1020[51][18]
        }
        else if(index==5){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][17]
          value.dopIdentifier=this.arrofFlowchartD1020[51][13]
          value.distance=this.arrofFlowchartD1020[51][19]
        }
        else if(index==6){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][12]
          value.dopIdentifier=this.arrofFlowchartD1020[52][13]
          value.distance=this.arrofFlowchartD1020[52][14]
        }
        else if(index==7){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][13]
          value.dopIdentifier=this.arrofFlowchartD1020[52][13]
          value.distance=this.arrofFlowchartD1020[52][15]
        }
        else if(index==8){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][14]
          value.dopIdentifier=this.arrofFlowchartD1020[52][13]
          value.distance=this.arrofFlowchartD1020[52][16]
        }
        else if(index==9){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][15]
          value.dopIdentifier=this.arrofFlowchartD1020[52][13]
          value.distance=this.arrofFlowchartD1020[52][17]
        }
        else if(index==10){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][16]
          value.dopIdentifier=this.arrofFlowchartD1020[52][13]
          value.distance=this.arrofFlowchartD1020[52][18]
        }
        else if(index==11){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][17]
          value.dopIdentifier=this.arrofFlowchartD1020[52][13]
          value.distance=this.arrofFlowchartD1020[52][19]
        }
        else if(index==12){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][12]
          value.dopIdentifier=this.arrofFlowchartD1020[53][13]
          value.distance=this.arrofFlowchartD1020[53][14]
        }
        else if(index==13){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][13]
          value.dopIdentifier=this.arrofFlowchartD1020[53][13]
          value.distance=this.arrofFlowchartD1020[53][15]
        }
        else if(index==14){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][14]
          value.dopIdentifier=this.arrofFlowchartD1020[53][13]
          value.distance=this.arrofFlowchartD1020[53][16]
        }
        else if(index==15){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][15]
          value.dopIdentifier=this.arrofFlowchartD1020[53][13]
          value.distance=this.arrofFlowchartD1020[53][17]
        }
        else if(index==16){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][16]
          value.dopIdentifier=this.arrofFlowchartD1020[53][13]
          value.distance=this.arrofFlowchartD1020[53][18]
        }
        else if(index==17){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][17]
          value.dopIdentifier=this.arrofFlowchartD1020[53][13]
          value.distance=this.arrofFlowchartD1020[53][19]
        }
        else if(index==18){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][12]
          value.dopIdentifier=this.arrofFlowchartD1020[54][13]
          value.distance=this.arrofFlowchartD1020[54][14]
        }
        else if(index==19){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][13]
          value.dopIdentifier=this.arrofFlowchartD1020[54][13]
          value.distance=this.arrofFlowchartD1020[54][15]
        }
        else if(index==20){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][14]
          value.dopIdentifier=this.arrofFlowchartD1020[54][13]
          value.distance=this.arrofFlowchartD1020[54][16]
        }
        else if(index==21){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][15]
          value.dopIdentifier=this.arrofFlowchartD1020[54][13]
          value.distance=this.arrofFlowchartD1020[54][17]
        }
        else if(index==22){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][16]
          value.dopIdentifier=this.arrofFlowchartD1020[54][13]
          value.distance=this.arrofFlowchartD1020[54][18]
        }
        else if(index==23){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][17]
          value.dopIdentifier=this.arrofFlowchartD1020[54][13]
          value.distance=this.arrofFlowchartD1020[54][19]
        }
        else if(index==24){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][12]
          value.dopIdentifier=this.arrofFlowchartD1020[55][13]
          value.distance=this.arrofFlowchartD1020[55][14]
        }
        else if(index==25){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][13]
          value.dopIdentifier=this.arrofFlowchartD1020[55][13]
          value.distance=this.arrofFlowchartD1020[55][15]
        }
        else if(index==26){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][14]
          value.dopIdentifier=this.arrofFlowchartD1020[55][13]
          value.distance=this.arrofFlowchartD1020[55][16]
        }
        else if(index==27){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][15]
          value.dopIdentifier=this.arrofFlowchartD1020[55][13]
          value.distance=this.arrofFlowchartD1020[55][17]
        }
        else if(index==28){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][16]
          value.dopIdentifier=this.arrofFlowchartD1020[55][13]
          value.distance=this.arrofFlowchartD1020[55][18]
        }
        else if(index==29){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][17]
          value.dopIdentifier=this.arrofFlowchartD1020[55][13]
          value.distance=this.arrofFlowchartD1020[55][19]
        }
        else if(index==30){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][12]
          value.dopIdentifier=this.arrofFlowchartD1020[56][13]
          value.distance=this.arrofFlowchartD1020[56][14]
        }
        else if(index==31){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][13]
          value.dopIdentifier=this.arrofFlowchartD1020[56][13]
          value.distance=this.arrofFlowchartD1020[56][15]
        }
        else if(index==32){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][14]
          value.dopIdentifier=this.arrofFlowchartD1020[56][13]
          value.distance=this.arrofFlowchartD1020[56][16]
        }
        else if(index==33){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][15]
          value.dopIdentifier=this.arrofFlowchartD1020[56][13]
          value.distance=this.arrofFlowchartD1020[56][17]
        }
        else if(index==34){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][16]
          value.dopIdentifier=this.arrofFlowchartD1020[56][13]
          value.distance=this.arrofFlowchartD1020[56][18]
        }
        else if(index==35){
          value.elevatorIdentifier=this.arrofFlowchartD1020[50][17]
          value.dopIdentifier=this.arrofFlowchartD1020[56][13]
          value.distance=this.arrofFlowchartD1020[56][19]
        }
      })
      this.walkingTimes.push(this.walkingTimeObj)
      this.visualGuidance.push(this.visualGuidanceObj)
      console.log("walkingTimes:",this.walkingTimes)
      console.log("visual Guidance:",this.visualGuidance)
      this.distance.distanceFromCore1ToCore2=this.arrofFlowchartD1020[112][1] || null
      this.distance.distanceFromCore1ToSiteController=this.arrofFlowchartD1020[113][1] || null
      this.distance.distanceFromCore2ToSiteController=this.arrofFlowchartD1020[114][1] || null
      this.distance.distanceFromCore1ToEdgeRouter1=this.arrofFlowchartD1020[115][1] || null
      this.distance.distanceFromCore2ToEdgeRouter1=this.arrofFlowchartD1020[116][1] || null
      this.distance.distanceFromCore2ToEdgeRouter2=this.arrofFlowchartD1020[117][1] || null
      this.distance.distanceFromCore1To1stGCM=this.arrofFlowchartD1020[118][1] || null
      this.distance.distanceFromCore2ToLastGCM=this.arrofFlowchartD1020[119][1] || null
      this.distance.distanceFromEdgeRouter1ToBuildingLAN=this.arrofFlowchartD1020[120][1] || null
      this.distance.distanceFromCore2ToFiberConverter2=this.arrofFlowchartD1020[123][1] || null
      this.cableDistance.cableFromGCMToController.G=this.arrofFlowchartD1020[125][1] || null
      this.cableDistance.cableFromGCMToController.H=this.arrofFlowchartD1020[125][2] || null
      this.cableDistance.cableFromGCMToController.I=this.arrofFlowchartD1020[125][3] || null
      this.cableDistance.cableFromGCMToController.J=this.arrofFlowchartD1020[125][4] || null
      this.cableDistance.cableFromGCMToController.K=this.arrofFlowchartD1020[125][5] || null
      this.cableDistance.cableFromGCMToController.L=this.arrofFlowchartD1020[125][6] || null
      this.cableDistance.cableFromGCMToServedFloorClosestToMachineRoom.G=this.arrofFlowchartD1020[126][1] || null
      this.cableDistance.cableFromGCMToServedFloorClosestToMachineRoom.H=this.arrofFlowchartD1020[126][2] || null
      this.cableDistance.cableFromGCMToServedFloorClosestToMachineRoom.I=this.arrofFlowchartD1020[126][3] || null
      this.cableDistance.cableFromGCMToServedFloorClosestToMachineRoom.J=this.arrofFlowchartD1020[126][4] || null
      this.cableDistance.cableFromGCMToServedFloorClosestToMachineRoom.K=this.arrofFlowchartD1020[126][5] || null
      this.cableDistance.cableFromGCMToServedFloorClosestToMachineRoom.L=this.arrofFlowchartD1020[126][6] || null
      this.cableDistance.cableFromGCMToGCM.G=this.arrofFlowchartD1020[128][1] || null
      this.cableDistance.cableFromGCMToGCM.H=this.arrofFlowchartD1020[128][2] || null
      this.cableDistance.cableFromGCMToGCM.I=this.arrofFlowchartD1020[128][3] || null
      this.cableDistance.cableFromGCMToGCM.J=this.arrofFlowchartD1020[128][4] || null
      this.cableDistance.cableFromGCMToGCM.K=this.arrofFlowchartD1020[128][5] || null
      this.cableDistance.cableFromGCMToGCM.L=this.arrofFlowchartD1020[128][6] || null
      this.cableDistance.CoreSwitchElevator.G=this.arrofFlowchartD1020[129][1] || null
      this.cableDistance.CoreSwitchElevator.H=this.arrofFlowchartD1020[129][2] || null
      this.cableDistance.CoreSwitchElevator.I=this.arrofFlowchartD1020[129][3] || null
      this.cableDistance.CoreSwitchElevator.J=this.arrofFlowchartD1020[129][4] || null
      this.cableDistance.CoreSwitchElevator.K=this.arrofFlowchartD1020[129][5] || null
      this.cableDistance.CoreSwitchElevator.L=this.arrofFlowchartD1020[129][6] || null
      this.cableDistance.LANRiserElevators.G=this.arrofFlowchartD1020[130][1] || null
      this.cableDistance.LANRiserElevators.H=this.arrofFlowchartD1020[130][2] || null
      this.cableDistance.LANRiserElevators.I=this.arrofFlowchartD1020[130][3] || null
      this.cableDistance.LANRiserElevators.J=this.arrofFlowchartD1020[130][4] || null
      this.cableDistance.LANRiserElevators.K=this.arrofFlowchartD1020[130][5] || null
      this.cableDistance.LANRiserElevators.L=this.arrofFlowchartD1020[130][6] || null
      this.cableDistance.cableFromCoreSwitchToFirstAccessSwitchInShaft.G=this.arrofFlowchartD1020[131][1] || null
      this.cableDistance.cableFromCoreSwitchToFirstAccessSwitchInShaft.H=this.arrofFlowchartD1020[131][2] || null
      this.cableDistance.cableFromCoreSwitchToFirstAccessSwitchInShaft.I=this.arrofFlowchartD1020[131][3] || null
      this.cableDistance.cableFromCoreSwitchToFirstAccessSwitchInShaft.J=this.arrofFlowchartD1020[131][4] || null
      this.cableDistance.cableFromCoreSwitchToFirstAccessSwitchInShaft.K=this.arrofFlowchartD1020[131][5] || null
      this.cableDistance.cableFromCoreSwitchToFirstAccessSwitchInShaft.L=this.arrofFlowchartD1020[131][6] || null
      this.cableDistance.distanceToBuildingOpticalpatchPanelOfIntegratedGroup.G=this.arrofFlowchartD1020[132][1] || null
      this.cableDistance.distanceToBuildingOpticalpatchPanelOfIntegratedGroup.H=this.arrofFlowchartD1020[132][2] || null
      this.cableDistance.distanceToBuildingOpticalpatchPanelOfIntegratedGroup.I=this.arrofFlowchartD1020[132][3] || null
      this.cableDistance.distanceToBuildingOpticalpatchPanelOfIntegratedGroup.J=this.arrofFlowchartD1020[132][4] || null
      this.cableDistance.distanceToBuildingOpticalpatchPanelOfIntegratedGroup.K=this.arrofFlowchartD1020[132][5] || null
      this.cableDistance.distanceToBuildingOpticalpatchPanelOfIntegratedGroup.L=this.arrofFlowchartD1020[132][6] || null
      this.cableDistance.siteControllerLocation.G=this.arrofFlowchartD1020[133][1] || null
      this.cableDistance.siteControllerLocation.H=this.arrofFlowchartD1020[133][2] || null
      this.cableDistance.siteControllerLocation.I=this.arrofFlowchartD1020[133][3] || null
      this.cableDistance.siteControllerLocation.J=this.arrofFlowchartD1020[133][4] || null
      this.cableDistance.siteControllerLocation.K=this.arrofFlowchartD1020[133][5] || null
      this.cableDistance.siteControllerLocation.L=this.arrofFlowchartD1020[133][6] || null
      this.cableDistance.cableFromAMSMToGCMAndIsolationSwitchToGCM.G=this.arrofFlowchartD1020[134][1] || null
      this.cableDistance.cableFromAMSMToGCMAndIsolationSwitchToGCM.H=this.arrofFlowchartD1020[134][2] || null
      this.cableDistance.cableFromAMSMToGCMAndIsolationSwitchToGCM.I=this.arrofFlowchartD1020[134][3] || null
      this.cableDistance.cableFromAMSMToGCMAndIsolationSwitchToGCM.J=this.arrofFlowchartD1020[134][4] || null
      this.cableDistance.cableFromAMSMToGCMAndIsolationSwitchToGCM.K=this.arrofFlowchartD1020[134][5] || null
      this.cableDistance.cableFromAMSMToGCMAndIsolationSwitchToGCM.L=this.arrofFlowchartD1020[134][6] || null
      this.cableDistance.cableFromAMSMToIsolationSwitch.G=this.arrofFlowchartD1020[135][1] || null
      this.cableDistance.cableFromAMSMToIsolationSwitch.H=this.arrofFlowchartD1020[135][2] || null
      this.cableDistance.cableFromAMSMToIsolationSwitch.I=this.arrofFlowchartD1020[135][3] || null
      this.cableDistance.cableFromAMSMToIsolationSwitch.J=this.arrofFlowchartD1020[135][4] || null
      this.cableDistance.cableFromAMSMToIsolationSwitch.K=this.arrofFlowchartD1020[135][5] || null
      this.cableDistance.cableFromAMSMToIsolationSwitch.L=this.arrofFlowchartD1020[135][6] || null
      this.cableDistance.AMSMToCoreSwitchandIsolationSwitchToCoreSwitch.G=this.arrofFlowchartD1020[136][1] || null
      this.cableDistance.AMSMToCoreSwitchandIsolationSwitchToCoreSwitch.H=this.arrofFlowchartD1020[136][2] || null
      this.cableDistance.AMSMToCoreSwitchandIsolationSwitchToCoreSwitch.I=this.arrofFlowchartD1020[136][3] || null
      this.cableDistance.AMSMToCoreSwitchandIsolationSwitchToCoreSwitch.J=this.arrofFlowchartD1020[136][4] || null
      this.cableDistance.AMSMToCoreSwitchandIsolationSwitchToCoreSwitch.K=this.arrofFlowchartD1020[136][5] || null
      this.cableDistance.AMSMToCoreSwitchandIsolationSwitchToCoreSwitch.L=this.arrofFlowchartD1020[136][6] || null
      this.cableDistance.extensionCableqty.G=this.arrofFlowchartD1020[137][1] || null
      this.cableDistance.extensionCableqty.H=this.arrofFlowchartD1020[137][2] || null
      this.cableDistance.extensionCableqty.I=this.arrofFlowchartD1020[137][3] || null
      this.cableDistance.extensionCableqty.J=this.arrofFlowchartD1020[137][4] || null
      this.cableDistance.extensionCableqty.K=this.arrofFlowchartD1020[137][5] || null
      this.cableDistance.extensionCableqty.L=this.arrofFlowchartD1020[137][6] || null
    
    } 
}
}
selectFile:any
onFileChanged(event:any) {
  const file:File = event.target.files[0];
  const formData=new FormData()
  formData.append('file',file)
  this.ex.postFile(formData).subscribe((value)=>{
    console.log(value)
  })
 
}

}
