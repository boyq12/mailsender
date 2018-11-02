import { Component, OnInit } from '@angular/core';
import * as ClassicEditor from '@ckeditor/ckeditor5-build-classic';
import { Http, Headers, RequestOptions } from "@angular/http";

import * as XLSX from 'xlsx';
type AOA = any[][];

var api_key = '30a5f6f102695b8249eb47a583023c24-0e6e8cad-b826fe4c';
var domain = 'sandbox345342effc71485aafbe4fff874301dd.mailgun.org';
const mailgunUrl: string = "MAILGUN_URL_HERE";
// const mailgun = require('mailgun-js')({ apiKey: api_key, domain: domain });



@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {

  public Editor = ClassicEditor;
  data: AOA = [];
  emailContent: string = "<p>Enter email here</p>";
  preEmail: string = "";
  settings = {
    columns: [],
    pager: {
      display: true,
      perPage: 10
    },
    actions: {
      add: false,
      edit: false,
      delete: false,
      position: 'left'
    },
    attr: {
      class: 'table table-striped table-bordered table-hover'
    },
    defaultStyle: false
  };
  convertedData = [];
  bulkEmail = [];
  sampleEmail = "";
  constructor(private http: Http) { }

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
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
      this.convertData();
    };
    reader.readAsBinaryString(target.files[0]);
  }

  convertData() {
    var newSettings = {
      columns: [],
      pager: {
        display: true,
        perPage: 10
      },
      actions: {
        add: false,
        edit: false,
        delete: false,
        position: 'left'
      },
      attr: {
        class: 'table table-striped table-bordered table-hover'
      },
      defaultStyle: false
    };
    this.data.forEach((values, i) => {
      if (i == 0) {
        values.forEach((header, pos) => {
          newSettings.columns[`${pos}`] = { title: header };
        });
      } else {
        var newRow = {};
        values.forEach((value, pos) => {
          newRow[`${pos}`] = value;
        });

        this.convertedData.push(newRow);
      }
    });
    this.settings = Object.assign({}, newSettings);
  }
  emailContentChange() {
    this.bulkEmail = Object.assign([], []);
    this.convertedData.forEach(data => {
      var tempEmail = { email: "", content: "" };
      tempEmail.content = this.emailContent;
      this.settings.columns.forEach((col, pos) => {
        var replaceCode = "${" + col.title + "}";
        if (col.title.toLowerCase() == "email") {
          tempEmail.email = data[pos];
        }
        tempEmail.content = tempEmail.content.replace(replaceCode, data[pos]);
      });
      this.bulkEmail.push(tempEmail);
    });
    console.log(this.bulkEmail);
    this.sampleEmail = this.bulkEmail[0].content || "";
  }

  sendEmails() {
    let headers = new Headers(
      {
        "Content-Type": "application/x-www-form-urlencoded",
        "Authorization": "Basic " + api_key
      }
    );
    let options = new RequestOptions({ headers: headers });
    this.bulkEmail.forEach(email => {
      let body = "from=cskh@easytrip.com&to=" + email.email + "&subject=" + "Hellp" + "&html=" + email.content;
      this.http.post("https://api.mailgun.net/v3/" + domain + "/messages", body, options)
                .subscribe(result => {
                  console.log(result)
                }, error => {
                    console.log(error);
                });
      // var data = {
      //   from: "Nam <leanhnam2203@gmail.com>",
      //   to: ["test@example.com"],
      //   subject: "Hello",
      //   html: email.content
      // };
      // mailgun.messages().send(data, { 'content-type': 'text/html' }, function (error, body) {
      //   console.log(body);
      // });
    })

  }

}
