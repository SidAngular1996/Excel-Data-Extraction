import { Component, OnInit } from '@angular/core';
import { FormBuilder, Validators, FormGroup } from '@angular/forms';
import * as XLSX from 'xlsx';
import { ArrayType } from '@angular/compiler';

type AOA = any[][];

@Component({
  selector: 'app-excel',
  templateUrl: './excel.component.html',
  styleUrls: ['./excel.component.css']
})
export class ExcelComponent implements OnInit {


  form: FormGroup;

  title = 'excel-sheets';


  Name: string = "Name*";
  Address1: string = "Address1*";
  Address2: string = "Address2*";
  City: string = "City*";
  State: string = "State*";
  PIN: string = "PIN*";
  ConsigneeName: string = "Consignee Name*";
  ConsigneeAddress1: string = "Consignee Address1*";
  ConsigneeAddress2: string = "Consignee Address2 ";


  data: AOA = [[this.Name, ''], [this.Address1, ''], [this.Address2, ''],
  [this.City, ''], [this.State, ''], [this.PIN, ''], [this.ConsigneeName, ''],
  [this.ConsigneeAddress1, ''], [this.ConsigneeAddress2]];

  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'SheetJS.xlsx';

  data_form: any

  Ename: String = ''
  Eadd1: String = ''
  Eadd2: String = ''
  Ecity: String = ''
  Estate: String = ''
  Epin: number
  Econname: String = ''
  Econadd1: String = ''
  Econadd2: String = ''

  //  Array_fields=[this.Ename,this.Eadd1,this.Eadd2,this.Ecity,this.Estate,this.Epin,this.Econname,this.Econadd1,this.Econadd2]
  Array_data = []

  constructor(private fb: FormBuilder) { }

  ngOnInit(): void {
    this.form = this.fb.group({
      Name: ['', Validators.required],
      Address1: ['', Validators.required],
      Address2: ['', Validators.required],
      City: ['', Validators.required],
      State: ['', Validators.required],
      PIN: ['', Validators.required],
      consigneeName: ['', Validators.required],
      consigneeAddress1: ['', Validators.required],
      consigneeAddress2: ['', Validators.required],
    })

  }
  onFileChange(evt: any) {
    const target = (evt.target);
    console.log("Target Files", target.files);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = function () {
      var data = reader.result
      console.log(data)
    }
    reader.onload = (e: any) => {
      console.log("E", e)
      // The result is a JavaScript ArrayBuffer containing binary data. The result contains the raw binary data from the file in a string. The result is a string with a data: URL representing the file's data. The result is text in a string
      const bstr: string = e.target.result;
      console.log("Bstr", bstr)

      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
      console.log("Wb", wb)

      const wsname: string = wb.SheetNames[0];
      console.log("Wsname", wsname)

      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      console.log("Ws", ws)

      this.data = (XLSX.utils.sheet_to_json(ws, { header: 1 }));
      this.data_form = this.data
      // this.Ename = this.data_form[0][1]
      this.data_fill()
      this.adjustingData()
    };

    reader.readAsBinaryString(target.files[0]);

    // this.data_form = this.data


  }

  export(): void {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const data_form = XLSX.utils.sheet_to_json(ws);
    console.log("Json", data_form)
    
    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }

  details() {
    console.log("Details Submitted")
  }

  data_fill() {

    var objectF = this.data_form

    console.log("Json Data", this.data_form[0][1]);

    console.log("objectF", objectF)

    objectF.forEach(element => {
      element = Object.assign({}, element)
      console.log(element)
      this.Array_data.push(element[1])

    });

    console.log("Array", this.Array_data)





  }
  adjustingData() {

    this.Ename = this.Array_data[0]
    this.Eadd1 = this.Array_data[1]
    this.Eadd2 = this.Array_data[2]
    this.Ecity = this.Array_data[3]
    this.Estate = this.Array_data[4]
    this.Epin = this.Array_data[5]
    this.Econname = this.Array_data[6]
    this.Econadd1 = this.Array_data[7]
    this.Econadd2 = this.Array_data[8]

  }
}
