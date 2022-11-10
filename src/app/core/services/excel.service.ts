import { Injectable } from '@angular/core';
import { Workbook, Worksheet } from 'exceljs';
import * as fs from 'file-saver';
import { Icompany } from '../interfaces/icompany';
import { LoggerService } from './logger.service';
@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  private IEmployeeDetails!: Icompany[];

  constructor(private logger: LoggerService) { }

  //Save Empty Content Excel

  createNewExcelFile(data: any) {

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
  private toDefineTitle(worksheet: Worksheet, title: string): Worksheet {

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
        color: { argb: 'FFFFFF' },
        size: 12
      };
    });

    worksheet.columns.forEach(function (column, i) {
      column.width = 30
    });

    return worksheet;
  }

  //load Uploaded File
  loadUploadedFile(importFile: any) {
    const workbook = new Workbook();
    const arrayBuffer = new Response(importFile).arrayBuffer();
    arrayBuffer.then((data) => {
      workbook.xlsx.load(data).then((workbook) => {
        if (this.validateUploadExcelTitle(workbook) && this.validateUploadExcelHeaders(workbook)) {
          workbook.getWorksheet(1).eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber > 5) {
              this.logger.logInformation('rowValue', row.values);
              this.logger.logInformation('rowNumber', rowNumber);
              this.IEmployeeDetails.push(row.values);
            }
          });
        }
        this.logger.logInformation('workbookDetails', workbook);
        this.logger.logInformation('workSheetDetails', workbook.getWorksheet(1));
        this.logger.logInformation('WorkSheetName', workbook.getWorksheet(1).name);
        this.logger.logInformation('WorkSheetTitle', workbook.getWorksheet(1).getCell('A2:F3').value);
        this.logger.logInformation('WorkHeaderRow', workbook.getWorksheet(1).getRow(5).values);
      });
    });
  }

  //validate Title Name
  private validateUploadExcelTitle(workbook: Workbook): boolean {
    return workbook.getWorksheet(1).getCell('A2:F3').value ? true : false;
  }

  //validate Header
  private validateUploadExcelHeaders(workbook: Workbook): boolean {
    const headers = ['EmployeeName', 'EmployeeSalary', 'EmployeeJoininDate', 'EmployeeRole', 'EmployeeEmail', 'EmployeePhone'];
    return (workbook.getWorksheet(1).getRow(5).values === headers) ? true : false;
  }
}
