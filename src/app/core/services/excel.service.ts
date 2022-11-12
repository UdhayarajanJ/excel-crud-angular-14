import { Injectable } from '@angular/core';
import { Workbook, Worksheet } from 'exceljs';
import * as fs from 'file-saver';
import { Observable } from 'rxjs';
import { Icompany } from '../interfaces/icompany';
import { Isheetdetails } from '../interfaces/isheetdetails';
import { LoggerService } from './logger.service';
@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  private IEmployeeDetails!: Icompany[];
  private IEmployeeDetailsObj!: Icompany;
  private IsheetDetails!: Isheetdetails;

  constructor(private logger: LoggerService) { }

  //Save Empty Content Excel

  public createNewExcelFile(data: any) {

    let workbook = new Workbook();

    //to define sheet Name
    let worksheet = workbook.addWorksheet(data.sheetName);

    //to define Title
    worksheet = this.toDefineTitle(worksheet, data.excelTitle);

    //to define Title
    worksheet = this.toDefineHeaders(worksheet);

    //to define FileName
    let fileName = data.fileName;

    //to Save The File
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/octet-stream' });
      fs.saveAs(blob, fileName + '.xlsx');
    });
  }

  //to define Title
  private toDefineTitle(worksheet: Worksheet, title: string | undefined): Worksheet {

    worksheet.mergeCells('A2:F3');
    worksheet.getCell('A2').value = title;

    worksheet.getRow(2).height = 25;
    worksheet.getRow(3).height = 25;

    worksheet.getCell('A2').alignment = {
      horizontal: 'center',
      vertical: 'middle'
    };

    worksheet.getCell('A2').style.fill = {
      pattern: 'solid',
      type: 'pattern',
      fgColor: { argb: 'ffdc3545' },
    };

    worksheet.getCell('A2').style.font = {
      bold: true,
      color: { argb: 'FFFFFF' },
      size: 24
    };
    return worksheet;
  }

  //to define FileName
  private toDefineHeaders(worksheet: Worksheet): Worksheet {

    const headers = ['EmployeeName', 'EmployeeSalary', 'EmployeeJoininDate', 'EmployeeRole', 'EmployeeEmail', 'EmployeePhone'];
    worksheet.insertRow(5, headers);

    worksheet.getRow(5).eachCell((cell, colNumber) => {
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle'
      }

      cell.style.fill = {
        pattern: 'solid',
        type: 'pattern',
        fgColor: { argb: 'ffdee2e6' },
      };

      cell.style.font = {
        bold: true,
        color: { argb: '000000' },
        size: 12
      };
    });

    worksheet.columns.forEach(function (column, i) {
      column.width = 30
    });

    return worksheet;
  }


  //validate Title Name
  public validateUploadExcelTitle(workbook: Workbook): boolean {
    //this.logger.logInformation('TitileValue', workbook.getWorksheet(1).getCell('A2:F3').value);
    let result = workbook.getWorksheet(1).getCell('A2:F3').value ? true : false;
    return result;
  }

  //validate Header
  public validateUploadExcelHeaders(workbook: Workbook): boolean {
    let headers = new Array();
    workbook.getWorksheet(1).getRow(5).eachCell({ includeEmpty: false }, (cell, cellNumber) => {
      headers.push(cell.value);
    });
    //this.logger.logInformation('HeadersValue', headers);
    const defaultHeader = ['EmployeeName', 'EmployeeSalary', 'EmployeeJoininDate', 'EmployeeRole', 'EmployeeEmail', 'EmployeePhone'];
    let result = JSON.stringify(headers) === JSON.stringify(defaultHeader) ? true : false;
    return result;
  }


  //To Get Sheet Details
  public toGetSheetDetails(workbook: Workbook, fileInfo: File, rowCount: number): Isheetdetails {
    this.IsheetDetails = {
      fileName: fileInfo.name,
      rowCount: rowCount,
      sheetName: workbook.getWorksheet(1).name,
      titleOfExcel: workbook.getWorksheet(1).getCell('A2:F3').value?.toString()
    }
    return this.IsheetDetails;
  }

  //To Save The File
  saveFile(sheetInfo: Isheetdetails, fileName: string, recordInfo: Icompany[]) {

    let workbook = new Workbook();

    let worksheet = workbook.addWorksheet(sheetInfo.sheetName);

    //to define Title
    worksheet = this.toDefineTitle(worksheet, sheetInfo.titleOfExcel);

    //to define Title
    worksheet = this.toDefineHeaders(worksheet);

    //to Insert Record
    recordInfo.forEach((value, index) => {
      const rowArray = [value.employeeName, value.employeeSalary, value.employeeJoininDate, value.employeeRole, value.employeeEmail, value.employeePhone];
      worksheet.insertRow(6 + index, rowArray);
      worksheet.getRow(6 + index).eachCell((cell, colNumber) => {
        cell.alignment = {
          horizontal: 'left',
          vertical: 'middle'
        }
        cell.style.font = {
          bold: true,
          color: { argb: '000000' },
          size: 10
        };
      });

    });

    //to Save The File
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/octet-stream' });
      fs.saveAs(blob, fileName + '.xlsx');
    });
  }
}
