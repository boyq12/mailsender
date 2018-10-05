import { Component, OnInit } from '@angular/core';
import * as ClassicEditor from '@ckeditor/ckeditor5-build-classic';

import * as XLSX from 'xlsx';
type AOA = any[][];

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {

  public Editor = ClassicEditor;
  data: AOA = [];
  settings = {
    columns : {}
  };
  convertedData = [];
  constructor() { }

  ngOnInit() {


  }

  onFileChange(evt: any) {
		/* wire up file reader */
		const target: DataTransfer = <DataTransfer>(evt.target);
		if (target.files.length !== 1) throw new Error('Cannot use multiple files');
		const reader: FileReader = new FileReader();
		reader.onload = (e: any) => {
			/* read workbook */
			const bstr: string = e.target.result;
			const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

			/* grab first sheet */
			const wsname: string = wb.SheetNames[0];
			const ws: XLSX.WorkSheet = wb.Sheets[wsname];

			/* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
      this.convertData();
		};
		reader.readAsBinaryString(target.files[0]);
  }
  
  convertData(){
    this.data.forEach((values, i) => {
      if(i == 0){
        values.forEach((header, pos) => {
          this.settings.columns[`${pos}`] = {title: header};
        });
      } else {
        var newRow = {};
        values.forEach((value, pos) => {
          newRow[`${pos}`] = value;
        });
        
        this.convertedData.push(newRow);
      }
    });
    console.log(this.settings);
    console.log(this.convertedData);
  }

}
