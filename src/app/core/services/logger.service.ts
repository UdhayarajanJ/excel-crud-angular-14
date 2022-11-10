import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';

@Injectable({
  providedIn: 'root',
})
export class LoggerService {

  private logger: boolean = environment.logger;

  constructor() { }

  logInformation(title: string, message: any) {
    this.logger ? console.log(title.concat('::::'), message) : '';
  }

  logWarning(title: string, message: any) {
    this.logger ? console.warn(title.concat('::::'), message) : '';
  }

  logError(title: string, message: any) {
    this.logger ? console.error(title.concat('::::'), message) : '';
  }
}
