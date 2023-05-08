import { Injectable, ElementRef } from '@angular/core';
import { Workbook } from 'exceljs';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
const EXCEL_EXTENSION = '.xlsx';
@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  fileName: string = 'BillingSheet.xlsx';
  constructor() {}

  public ExcelTable(json: any[], excelFileName: string): void {
    const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
    XLSX.writeFile(workbook, this.fileName);
  }

  public ExportAsExcel(
    reportheading: string,
    headersArray: any[],
    json: any[],
    excelFileName: string,
    filename: string
  ) {
    const data = json;
    const header = headersArray;

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(filename);

    worksheet.addRow([]);
    worksheet.getCell('A1').value = reportheading;
    worksheet.getCell('A1').alignment = { horizontal: 'center' };
    worksheet.getCell('A1').font = { size: 14, bold: true };

    worksheet.addRow([]);

    const heading = worksheet.addRow(header);
    heading.eachCell((cell: any) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' },
        bgColor: { argb: 'FFFF00' },
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
      cell.font = { size: 12, bold: false };
    });

    data.forEach((element: any) => {
      const row = worksheet.addRow(element);
      console.log('row', row);
    });

    worksheet.addRow([]);
    worksheet.getColumn(1).width = 30;
    worksheet.properties.defaultColWidth = 17;

    workbook.xlsx.writeBuffer().then((data: ArrayBuffer) => {
      const blob = new Blob([data]);
      console.log('eachRowe', data);
      FileSaver.saveAs(blob, excelFileName + EXCEL_EXTENSION);
    });
  }
}
