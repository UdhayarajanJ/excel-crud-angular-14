import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { OnlynumberDirective } from './directives/onlynumber.directive';
import { OnlycharDirective } from './directives/onlychar.directive';

const sharedDirectives = [
  OnlynumberDirective,
  OnlycharDirective
];


@NgModule({
  declarations: [
    sharedDirectives
  ],
  imports: [
    CommonModule
  ],
  exports: [
    sharedDirectives
  ]
})
export class SharedModule { }
