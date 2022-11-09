import { Component, ElementRef, OnInit, TemplateRef, ViewChild } from '@angular/core';
import { NgxSpinnerService } from 'ngx-spinner';
import { ToastrService } from 'ngx-toastr';
import { LoggerService } from './core/services/logger.service';

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

  //Declare Hide Show Button And Radio 
  @ViewChild('optionCheckedCreateFile') optionCheckedCreateFile!: ElementRef;
  @ViewChild('optionCheckedUploadFile') optionCheckedUploadFile!: ElementRef;

  constructor(
    private spinner: NgxSpinnerService,
    private logger: LoggerService,
    private toaster: ToastrService
  ) {

  }

  ngOnInit(): void {
    this.logger.logInformation('CheckLogger', 'Angular 14 Excel Crud');
    //this.toaster.success('Success','Message');
  }

  optionsChecked(e: any) {
    if (e.target.value == 1) {
      this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Download_File_Popup');
      this.openModalPopup.nativeElement.click();
    }
    else {
      this.openModalPopup.nativeElement.setAttribute('data-bs-target', '#btn_Upload_File_Popup');
      this.openModalPopup.nativeElement.click();
    }
  }

  popupCancelSelection() {
    this.optionCheckedCreateFile.nativeElement.checked=false;
    this.optionCheckedUploadFile.nativeElement.checked=false;
  }
}
