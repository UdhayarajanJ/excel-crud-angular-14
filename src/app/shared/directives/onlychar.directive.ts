import { Directive, ElementRef, HostListener } from '@angular/core';

@Directive({
  selector: '[appOnlychar]'
})
export class OnlycharDirective {

  regexString: string = '^[A-Za-z ]+$';

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
    const pastData = event.clipboardData?.getData('text/plain').replace(/[^A-Za-z ]/g, '');
    this.el.nativeElement.value = pastData;
  }
}
