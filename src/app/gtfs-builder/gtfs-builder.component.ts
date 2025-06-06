import { Component } from '@angular/core';
import { MatButtonModule } from '@angular/material/button';
import { MatStepperModule } from '@angular/material/stepper';
import { ExcelUtilsService } from '../services/excel-utils.service';

@Component({
  selector: 'app-gtfs-builder',
  imports: [MatButtonModule, MatStepperModule],
  templateUrl: './gtfs-builder.component.html',
  styleUrl: './gtfs-builder.component.scss',
})
export class GtfsBuilderComponent {
  constructor(private excelUtilsService: ExcelUtilsService) {}

  onFileSelected(event: any) {
    console.log(event.target.files);
    const files = event.target.files;
    for (const file of files) {
      const name: string = file.name;
      const format: string | undefined = name.split('.').pop();
      console.log(name, file.size);
      console.log(file.size);
      if (format === 'xlsx') {
        console.log('Excel file');
        this.excelUtilsService.readFile(file);
      } else {
        console.log('Not an Excel file');
      }
    }
  }
}
