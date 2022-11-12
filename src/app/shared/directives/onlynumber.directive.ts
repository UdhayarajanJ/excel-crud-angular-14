import { Directive, ElementRef, HostListener } from '@angular/core';

@Directive({
  selector: '[appOnlynumber]'
})
export class OnlynumberDirective {

  regexString: string = '^[0-9]+$';

  constructor(private el: ElementRef) { }

  @HostListener('keypress', ['$event'])
  onKeyPress(event: any)
  {
    return new RegExp(this.regexString).test(event.key);
  }

  @HostListener('input', ['$event'])
  onInput(event: any)
  {
    return new RegExp(this.regexString).test(event.key);
  }

  @HostListener('paste', ['$event'])
  onPaste(event: ClipboardEvent)
  {
    this.validateFields(event);
  }

  validateFields(event: ClipboardEvent)
  {
    event.preventDefault();
    const pastData = event.clipboardData?.getData('text/plain').replace(/[^0-9]/g, '');
    this.el.nativeElement.value = pastData;
  }

}
