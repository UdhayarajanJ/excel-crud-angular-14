import { ThisReceiver } from '@angular/compiler';
import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { FormBuilder, FormGroup, Validators } from '@angular/forms';
import { NgxSpinnerService } from 'ngx-spinner';
import { ToastrService } from 'ngx-toastr';
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

  createFileFormGroup!: FormGroup;
  isCreateFileFormIsSubmitted: boolean = false;
  constructor(
    private spinner: NgxSpinnerService,
    private logger: LoggerService,
    private toaster: ToastrService,
    private fb: FormBuilder
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
      this.logger.logInformation('Form Value', this.createFileFormGroup.value);
    }
  }
}
