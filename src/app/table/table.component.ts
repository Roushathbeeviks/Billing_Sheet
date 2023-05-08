import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';
import DataJson from '../../app/data.json';
import { ExcelService } from '../excel.service';
import * as XLSX from 'xlsx';

interface DATA {
  name: String;
  day1?: string;
  day2?: string;
  day3?: string;
  day4?: string;
  day5?: string;
  day6?: string;
  day7?: string;
  day8?: string;
  day9?: string;
  day10?: string;
  day11?: string;
  day12?: string;
  day13?: string;
  day14?: string;
  day15?: string;
  day16?: string;
  day17?: string;
  day18?: string;
  day19?: string;
  day20?: string;
  day21?: string;
  day22?: string;
  day23?: string;
  day24?: string;
  day25?: string;
  day26?: string;
  day27?: string;
  day28?: string;
  day29?: string;
  day30?: string;
}

@Component({
  selector: 'app-table',
  templateUrl: './table.component.html',
  styleUrls: ['./table.component.css'],
})
export class TableComponent implements OnInit {
  @ViewChild('Excel') Excel: ElementRef;

  Data: DATA[] = DataJson;
  orderheader: string = '';
  search: string = '';

  order: string = 'name';
  reverse: boolean = false;

  fileName: string = 'BillingSheet.xlsx';

  p: number = 1;

  array: any;
  view: boolean = false;
  rowCount: any = 0;
  public tableData: any;
  public tableTitle: any;
  public tableSubTitle: any;
  public tableRecords = [];
  public pageStartCount = 1; //first row is header itself
  public pageEndCount: any; //
  public totalPageCount = 0;
  public columns: any;
  public xlColumns = [
    'Name',
    'Day-1',
    'Day-2',
    'Day-3',
    'Day-4',
    'Day-5',
    'Day-6',
    'Day-7',
    'Day-8',
    'Day-9',
    'Day-10',
    'Day-11',
    'Day-12',
    'Day-13',
    'Day-14',
    'Day-15',
    'Day-16',
    'Day-17',
    'Day-18',
    'Day-19',
    'Day-20',
    'Day-21',
    'Day-22',
    'Day-23',
    'Day-24',
    'Day-25',
    'Day-26',
    'Day-27',
    'Day-28',
    'Day-29',
    'Day-30',
  ];

  constructor(private ExcelService: ExcelService) {}

  ngOnInit(): void {}

  SortBy(name: string) {
    if (this.order === name) {
      this.reverse = !this.reverse;
    }
    this.order = name;
  }

  Download() {
    let element = document.getElementById('excel');
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(element);
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, this.fileName);
  }

  Export() {
    return new Promise((resolve) => {
      const sortedList: any = [];
      this.Data.forEach((data) => {
        const newObj: any = {};
        this.xlColumns.forEach((column) => {
          // console.log(column, 'column');
          switch (column) {
            case 'Name':
              newObj[column] = data?.name;
              break;
            case 'Day-1':
              newObj[column] = data?.day1;
              break;
            case 'Day-2':
              newObj[column] = data?.day2;
              break;
            case 'Day-3':
              newObj[column] = data?.day3;
              break;
            case 'Day-4':
              newObj[column] = data?.day4;
              break;
            case 'Day-5':
              newObj[column] = data?.day5;
              break;
            case 'Day-6':
              newObj[column] = data?.day6;
              break;
            case 'Day-7':
              newObj[column] = data?.day7;
              break;
            case 'Day-8':
              newObj[column] = data?.day8;
              break;
            case 'Day-9':
              newObj[column] = data?.day9;
              break;
            case 'Day-10':
              newObj[column] = data?.day10;
              break;
            case 'Day-11':
              newObj[column] = data?.day11;
              break;
            case 'Day-12':
              newObj[column] = data?.day12;
              break;
            case 'Day-13':
              newObj[column] = data?.day13;
              break;
            case 'Day-14':
              newObj[column] = data?.day14;
              break;
            case 'Day-15':
              newObj[column] = data?.day15;
              break;
            case 'Day-16':
              newObj[column] = data?.day16;
              break;
            case 'Day-17':
              newObj[column] = data?.day17;
              break;
            case 'Day-18':
              newObj[column] = data?.day18;
              break;
            case 'Day-19':
              newObj[column] = data?.day19;
              break;
            case 'Day-20':
              newObj[column] = data?.day20;
              break;
            case 'Day-21':
              newObj[column] = data?.day21;
              break;
            case 'Day-22':
              newObj[column] = data?.day22;
              break;
            case 'Day-23':
              newObj[column] = data?.day23;
              break;
            case 'Day-24':
              newObj[column] = data?.day24;
              break;
            case 'Day-25':
              newObj[column] = data?.day25;
              break;
            case 'Day-26':
              newObj[column] = data?.day26;
              break;
            case 'Day-27':
              newObj[column] = data?.day27;
              break;
            case 'Day-28':
              newObj[column] = data?.day28;
              break;
            case 'Day-29':
              newObj[column] = data?.day29;
              break;
            case 'Day-30':
              newObj[column] = data?.day30;
              break;
          }
        });
        sortedList.push(newObj);
        console.log('sortlist', sortedList);
        this.array = sortedList.map((obj: any) => Object.values(obj)); //Convert the json(array of objects) to array
        console.log('array', this.array);
      });

      this.ExcelService.ExportAsExcel(
        'BILLING SHEET',
        this.xlColumns,
        this.array,
        'BillingSheet',
        'Billingsheet'
      );
      //  this.ExcelService.ExcelTable(this.array,'Excel')
    });
  }

  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>evt.target;
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
      const data = XLSX.utils.sheet_to_json(ws);
      console.log('data1:', data);

      this.tableData = data;
      //  First Row is removed (Heading)
      console.log('this.tableData[1]', this.tableData[0]);
      this.tableTitle = Object.keys(this.tableData[0]);
      this.tableSubTitle = Object.values(this.tableData[0]);
      console.log('this.tableTitle:', this.tableTitle);
      console.log('this.tableSubTitle:', this.tableSubTitle);
      this.tableRecords = this.tableData.slice(
        this.pageStartCount,
        this.pageEndCount
      );
      console.log('this.tableData.slice', this.tableData.slice);
      console.log('this.tableRecords:', this.tableRecords);
    };
    reader.readAsBinaryString(target.files[0]);
    this.view = true;
  }
}
