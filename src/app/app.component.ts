import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { Workbook } from 'exceljs';
import { NgxSpinnerService } from 'ngx-spinner';
import { ToastrService } from 'ngx-toastr';
import { Icompany } from './core/interfaces/icompany';
import { Isheetdetails } from './core/interfaces/isheetdetails';
import { ExcelService } from './core/services/excel.service';
import { LoggerService } from './core/services/logger.service';
import { from, of } from 'rxjs';
import { skip, filter, take } from 'rxjs/operators';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {

  //Declare Template Reference Variable
  @ViewChild('openModalPopup') openModalPopup!: ElementRef;
  @ViewChild('closeUploadFilePopup') closeUploadFilePopup!: ElementRef;
  @ViewChild('closeCreateFilePopup') closeCreateFilePopup!: ElementRef;
  @ViewChild('closeAddEditModalPopup') closeAddEditModalPopup!: ElementRef;
  @ViewChild('closeDeleteRecordPopup') closeDeleteRecordPopup!: ElementRef;
  @ViewChild('closeSaveExcelPopup') closeSaveExcelPopup!: ElementRef;

  //Declare Hide Show Button And Radio 
  @ViewChild('optionCheckedCreateFile') optionCheckedCreateFile!: ElementRef;
  @ViewChild('optionCheckedUploadFile') optionCheckedUploadFile!: ElementRef;

  //Excel Row Information And Sheet information
  public IEmployeeDetails!: Icompany[];
  private IEmployeeDetailsObj!: Icompany;
  public IsheetDetails!: Isheetdetails;

  //Declare File Selection Variable
  selectedFile!: File;

  //Delare Form Regarding Variables
  createFileFormGroup!: FormGroup;
  isCreateFileFormIsSubmitted: boolean = false;
  addNewRecordForm!: FormGroup;
  modalHeaderForAdd_Edit_Record: string = 'Add Record';
  buttonForAdd_Edit_Record: string = 'Add';
  isAddEditFormIsSubmitted: boolean = false;
  isUpdated: boolean = false;
  indexNumber: number = -1;
  invalidFileName: boolean = false;
  isChangedFileTrue: boolean = false;
  showTableView: boolean = false;
  arrayCompany: any;

  //Pagination Variable
  pageNo: number = 1;
  pageSize: number = 3;
  offsetValue: number = 0;


  constructor(
    private spinner: NgxSpinnerService,
    private logger: LoggerService,
    private toaster: ToastrService,
    private fb: FormBuilder,
    private excel: ExcelService
  ) {

  }

  ngOnInit(): void {
    this.logger.logInformation('CheckLogger', 'Angular 14 Excel Crud');
    this.defaultLoadForm();
    //this.toaster.success('Success','Message');
  }

  //Option Changes Then Shown Modal Popup
  optionsChecked(e: any) {
    if (e.target.value == 1) {
      this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Download_File_Popup');
      this.openModalPopup.nativeElement.click();
      this.showTableView = false;
    }
    else {
      this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Upload_File_Popup');
      this.openModalPopup.nativeElement.click();
    }
  }

  //Close Popup To Unchecked The Option
  popupCancelSelection() {
    this.optionCheckedCreateFile.nativeElement.checked = false;
    this.optionCheckedUploadFile.nativeElement.checked = false;
    this.clearFormData();

  }

  //Form Loading 
  defaultLoadForm() {
    this.createFileFormGroup = this.fb.group({
      excelTitle: ['', [Validators.required]],
      sheetName: ['', [Validators.required]],
      fileName: ['', [Validators.required]],
    });

    this.addNewRecordForm = this.fb.group({
      employeeName: ['', [Validators.required]],
      employeeSalary: ['', [Validators.required]],
      employeeJoininDate: ['', [Validators.required]],
      employeeRole: ['', [Validators.required]],
      employeeEmail: ['', [Validators.required, Validators.pattern('[a-z0-9._%+-]+@[a-z0-9.-]+\\.[a-z]{2,4}$')]],
      employeePhone: ['', [Validators.required, Validators.pattern('^[6-9]{1}[0-9]{9}$')]]
    });
  }

  //Get Form Errors 
  get f() {
    return this.createFileFormGroup.controls;
  }

  get fAddEditForm() {
    return this.addNewRecordForm.controls;
  }

  //Create File On Submit
  onSubmitCreateFile() {
    this.logger.logInformation('Submit Check', 'Submit Working Fine');
    this.isCreateFileFormIsSubmitted = true;
    if (this.createFileFormGroup.invalid)
      return;
    else {
      this.spinner.show();
      this.excel.createNewExcelFile(this.createFileFormGroup.value);
      this.closeCreateFilePopup.nativeElement.click();
      this.toaster.success('File Created Successfully...', 'Message');
      this.spinner.hide();
      this.logger.logInformation('Form Value', this.createFileFormGroup.value);
    }
  }


  //onChange File
  onChange(e: any) {
    this.selectedFile = e.target.files[0];
    const extension = this.selectedFile.name.substring(this.selectedFile.name.lastIndexOf('.') + 1, this.selectedFile.name.length);
    //this.logger.logInformation('checkExtension', extension);
    if (extension == 'xlsx') {
      this.loadUploadedFile(this.selectedFile);
    }
    else {
      this.toaster.error('Invalid File Format...', 'Message');
    }
  }


  public loadUploadedFile(importFile: any): void {
    let result = {};
    const workbook = new Workbook();
    const arrayBuffer = new Response(importFile).arrayBuffer();
    arrayBuffer.then((data) => {
      workbook.xlsx.load(data).then((workbook) => {
        if (this.excel.validateUploadExcelTitle(workbook) && this.excel.validateUploadExcelHeaders(workbook)) {
          let employeeDetails = new Array();

          //Read Row Of Excel Sheet
          workbook.getWorksheet(1).eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber > 5) {
              let rowsValue = new Array();
              workbook.getWorksheet(1).getRow(rowNumber).eachCell({ includeEmpty: false }, (cell, cellNumber) => {
                rowsValue.push(cell.value)
              });
              this.IEmployeeDetailsObj = {
                employeeName: rowsValue[0],
                employeeSalary: rowsValue[1],
                employeeJoininDate: rowsValue[2],
                employeeRole: rowsValue[3],
                employeeEmail: rowsValue[4].text ? rowsValue[4].text : rowsValue[4],
                employeePhone: rowsValue[5]
              };
              employeeDetails.push(this.IEmployeeDetailsObj);

              this.logger.logInformation('rowInformation', this.IEmployeeDetails);
              this.logger.logInformation('sheetInformation', this.IsheetDetails);
            }

          });
          //Read Excel Row Info
          this.IEmployeeDetails = employeeDetails;
          this.pageNo = 1;
          this.paginateArrayOfCompany();
          //Read Sheet Details Info
          this.IsheetDetails = this.excel.toGetSheetDetails(workbook, importFile, this.IEmployeeDetails.length);
          this.showTableView = true;
          this.closeUploadFilePopup.nativeElement.click();
          this.toaster.success('File Upload Successfully...', 'Message');
        }
        else {
          this.toaster.error('Invalid File Data Please Enter Correct Structured File', 'Message');
        }
      });
    });
  }

  //Add Edit Submitted Form On Submit
  onSubmitAddEditRecord() {
    this.logger.logInformation('Submit Check', 'Submit Working Fine');
    this.isAddEditFormIsSubmitted = true;
    if (this.addNewRecordForm.invalid)
      return;
    const formValue = this.addNewRecordForm.value;
    this.IEmployeeDetailsObj = {
      employeeName: formValue.employeeName,
      employeeSalary: formValue.employeeSalary,
      employeeJoininDate: formValue.employeeJoininDate,
      employeeRole: formValue.employeeRole,
      employeeEmail: formValue.employeeEmail,
      employeePhone: formValue.employeePhone
    };
    if (!this.isUpdated) {
      this.IEmployeeDetails.push(this.IEmployeeDetailsObj);
      this.IsheetDetails.rowCount = this.IEmployeeDetails.length;
      this.pageNo = 1;
      this.paginateArrayOfCompany();
      this.toaster.success('Added New Record', 'Message');
    }
    else {
      this.IEmployeeDetails.splice(this.indexNumber, 1, this.IEmployeeDetailsObj);
      this.IsheetDetails.rowCount = this.IEmployeeDetails.length;
      this.toaster.success('Record Updated...', 'Message');
    }
    this.isChangedFileTrue = true;
    this.closeAddEditModalPopup.nativeElement.click();
    this.logger.logInformation('Form Value', this.addNewRecordForm.value);

  }

  clearFormData() {
    this.defaultLoadForm();
    this.isCreateFileFormIsSubmitted = false;
    this.isAddEditFormIsSubmitted = false;
    this.modalHeaderForAdd_Edit_Record = 'Add Record';
    this.buttonForAdd_Edit_Record = 'Add';
    this.isUpdated = false;
    this.indexNumber = -1;
  }

  editRecord(icompany: Icompany, index: number) {
    this.isUpdated = true;
    this.modalHeaderForAdd_Edit_Record = 'Update Record';
    this.buttonForAdd_Edit_Record = 'Update';
    this.addNewRecordForm.patchValue({
      employeeName: icompany.employeeName,
      employeeSalary: icompany.employeeSalary,
      employeeJoininDate: icompany.employeeJoininDate,
      employeeRole: icompany.employeeRole,
      employeeEmail: icompany.employeeEmail,
      employeePhone: icompany.employeePhone,
    });
    this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Add_Edit_Record');
    this.openModalPopup.nativeElement.click();
    this.indexNumber = index;
  }

  deleteRecord() {
    this.IEmployeeDetails.splice(this.indexNumber, 1);
    this.IsheetDetails.rowCount = this.IEmployeeDetails.length;
    this.closeDeleteRecordPopup.nativeElement.click();
    this.isChangedFileTrue = true;
    this.pageNo = 1;
    this.paginateArrayOfCompany();
    this.toaster.success('Record Deleted...', 'Message');
  }

  saveExcelFile(fileName: any) {
    this.logger.logInformation('fileName Txt', fileName.value);
    if (fileName.value == null || fileName.value == undefined || fileName.value == '') {
      this.invalidFileName = true;
      return;
    }
    this.closeSaveExcelPopup.nativeElement.click();
    this.excel.saveFile(this.IsheetDetails, fileName.value, this.IEmployeeDetails);
    this.toaster.success('New File Saved...', 'Message');
    fileName.value = '';
    this.isChangedFileTrue = false;
    this.showTableView = false;
  }

  onPageChangeEvent(event: any) {
    this.pageNo = event;
    this.paginateArrayOfCompany();
  }

  paginateArrayOfCompany() {
    this.offsetValue = (this.pageNo - 1) * this.pageSize;
    const icompanyInfo = of(this.IEmployeeDetails);
    const paginationBase = icompanyInfo.pipe(skip(this.offsetValue), take(this.pageSize));
    paginationBase.subscribe({
      next: (data) => {
        this.arrayCompany = data;
      },
      error: (err) => {
        this.logger.logError('paginateError', err);
      }
    });
  }
}
