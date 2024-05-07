export interface consumerValue {
    value: any;
    error?: any;
    id?:string;
  }
 export interface consumerForm {
    pageNo:number;
    carryDtToSqntPage: consumerValue;
    startNewDocument: consumerValue;
    consumer: consumerValue;
    documentType: consumerValue;
    documentName: consumerValue;
    documentDesc: consumerValue;
    image?:string;
  } 
  export enum HeaderButtonTypes{
    InitiateScanning,
    loadDiskDocument,
    ScanningCompleted
  }