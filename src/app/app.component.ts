import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { Workbook } from 'exceljs';
import { NgxSpinnerService } from 'ngx-spinner';
import { ToastrService } from 'ngx-toastr';
import { Icompany } from './core/interfaces/icompany';
import { Isheetdetails } from './core/interfaces/isheetdetails';
import { ExcelService } from './core/services/excel.service';
import { LoggerService } from './core/services/logger.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit
{

  //Declare Template Reference Variable
  @ViewChild('openModalPopup') openModalPopup!: ElementRef;
  @ViewChild('closeUploadFilePopup') closeUploadFilePopup!: ElementRef;
  @ViewChild('closeCreateFilePopup') closeCreateFilePopup!: ElementRef;

  //Declare Hide Show Button And Radio 
  @ViewChild('optionCheckedCreateFile') optionCheckedCreateFile!: ElementRef;
  @ViewChild('optionCheckedUploadFile') optionCheckedUploadFile!: ElementRef;

  //Excel Row Information And Sheet information
  public IEmployeeDetails!: Icompany[];
  private IEmployeeDetailsObj!: Icompany;
  public IsheetDetails!: Isheetdetails;

  //Declare File Selection Variable
  selectedFile!: File;

  createFileFormGroup!: FormGroup;
  isCreateFileFormIsSubmitted: boolean = false;
  constructor(
    private spinner: NgxSpinnerService,
    private logger: LoggerService,
    private toaster: ToastrService,
    private fb: FormBuilder,
    private excel: ExcelService
  )
  {

  }

  ngOnInit(): void
  {
    this.logger.logInformation('CheckLogger', 'Angular 14 Excel Crud');
    this.defaultLoadForm();
    //this.toaster.success('Success','Message');
  }

  //Option Changes Then Shown Modal Popup
  optionsChecked(e: any)
  {
    if (e.target.value == 1)
    {
      this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Download_File_Popup');
      this.openModalPopup.nativeElement.click();
    }
    else
    {
      this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Upload_File_Popup');
      this.openModalPopup.nativeElement.click();
    }
  }

  //Close Popup To Unchecked The Option
  popupCancelSelection()
  {
    this.optionCheckedCreateFile.nativeElement.checked = false;
    this.optionCheckedUploadFile.nativeElement.checked = false;
    this.defaultLoadForm();
    this.isCreateFileFormIsSubmitted = false;
  }

  //Form Loading 
  defaultLoadForm()
  {
    this.createFileFormGroup = this.fb.group({
      excelTitle: ['', [Validators.required]],
      sheetName: ['', [Validators.required]],
      fileName: ['', [Validators.required]],
    });
  }

  //Get Form Errors 
  get f()
  {
    return this.createFileFormGroup.controls;
  }

  //Create File On Submit
  onSubmitCreateFile()
  {
    this.logger.logInformation('Submit Check', 'Submit Working Fine');
    this.isCreateFileFormIsSubmitted = true;
    if (this.createFileFormGroup.invalid)
      return;
    else
    {
      this.spinner.show();
      this.excel.createNewExcelFile(this.createFileFormGroup.value);
      this.closeCreateFilePopup.nativeElement.click();
      this.toaster.success('File Created Successfully...', 'Message');
      this.spinner.hide();
      this.logger.logInformation('Form Value', this.createFileFormGroup.value);
    }
  }


  //onChange File
  onChange(e: any)
  {
    this.selectedFile = e.target.files[0];
    const extension = this.selectedFile.name.substring(this.selectedFile.name.lastIndexOf('.') + 1, this.selectedFile.name.length);
    //this.logger.logInformation('checkExtension', extension);
    if (extension == 'xlsx')
    {
      this.loadUploadedFile(this.selectedFile);
    }
    else
    {
      this.toaster.error('Invalid File Format...', 'Message');
    }
  }


  public loadUploadedFile(importFile: any): void
  {
    let result = {};
    const workbook = new Workbook();
    const arrayBuffer = new Response(importFile).arrayBuffer();
    arrayBuffer.then((data) =>
    {
      workbook.xlsx.load(data).then((workbook) =>
      {
        if (this.excel.validateUploadExcelTitle(workbook) && this.excel.validateUploadExcelHeaders(workbook))
        {
          let employeeDetails = new Array();

          //Read Row Of Excel Sheet
          workbook.getWorksheet(1).eachRow({ includeEmpty: false }, (row, rowNumber) =>
          {
            if (rowNumber > 5)
            {
              let rowsValue = new Array();
              workbook.getWorksheet(1).getRow(rowNumber).eachCell({ includeEmpty: false }, (cell, cellNumber) =>
              {
                rowsValue.push(cell.value)
              });
              this.IEmployeeDetailsObj = {
                employeeName: rowsValue[0],
                employeeSalary: rowsValue[1],
                employeeJoininDate: rowsValue[2],
                employeeRole: rowsValue[3],
                employeeEmail: rowsValue[4].text,
                employeePhone: rowsValue[5]
              };
              employeeDetails.push(this.IEmployeeDetailsObj);
              this.logger.logInformation('rowInformation', this.IEmployeeDetails);
              this.logger.logInformation('sheetInformation', this.IsheetDetails);
            }

          });
          //Read Excel Row Info
          this.IEmployeeDetails = employeeDetails;
          //Read Sheet Details Info
          this.IsheetDetails = this.excel.toGetSheetDetails(workbook, importFile, this.IEmployeeDetails.length);
          this.closeUploadFilePopup.nativeElement.click();
          this.toaster.success('File Upload Successfully...', 'Message');
        }
        else
        {
          this.toaster.error('Invalid File Data Please Enter Correct Structured File', 'Message');
        }
      });
    });
  }
}
