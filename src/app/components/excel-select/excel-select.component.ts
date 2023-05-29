import * as XLSX from 'xlsx';
import { Component } from '@angular/core';

@Component({
  selector: 'component-excel-select',
  templateUrl: './excel-select.component.html',
  styles: [],
})
export class ExcelSelectComponent {
  excelData: any;
  celdaSeleccionada: string = '';
  valorCeldaSeleccionada: string | undefined;

  ReadExcel(event: any) {
    let file = event.target.files[0];

    let fileReader = new FileReader();

    fileReader.readAsBinaryString(file);

    fileReader.onload = (e) => {
      let workBook = XLSX.read(fileReader.result, { type: 'binary' });
      //? console.log(workBook);

      var sheetNames = workBook.SheetNames;
      // sheetNames es una lista con los nombres de todas las hojas
      // i.e ['base','hoja2', ...]
      //? console.log(sheetNames);

      let workSheet = workBook.Sheets[sheetNames[0]];
      // workSheet contiene toda la info de la hoja en cuestion
      // ? console.log(workSheet);

      let cell = workSheet['B2'];
      console.log(cell);
      let cellValue = cell ? cell.v : undefined;
      console.log(cellValue);

      this.excelData = XLSX.utils.sheet_to_json(workBook.Sheets[sheetNames[0]]);
      console.log(this.excelData);
    };
  }

  // B2 C3 D4
  testFunc(event: any) {
    let file = event.target.files[0];
    let fileReader = new FileReader();
    fileReader.readAsBinaryString(file);

    fileReader.onload = (e) => {
      let workBook = XLSX.read(fileReader.result, { type: 'binary' });
      //? console.log(workBook);

      var sheetNames = workBook.SheetNames;
      // sheetNames es una lista con los nombres de todas las hojas
      // i.e ['base','hoja2', ...]
      //? console.log(sheetNames);

      let workSheet = workBook.Sheets[sheetNames[0]];
      // workSheet contiene toda la info de la hoja en cuestion
      console.log(workSheet);

      let cell = workSheet[this.celdaSeleccionada];
      //? console.log(cell);
      let cellValue = cell ? cell.v : undefined;
      this.valorCeldaSeleccionada = cellValue;
      //? console.log(cellValue);
    };
  }

  cambiarCelda(valor: string) {
    this.celdaSeleccionada = valor;
  }

  reiniciar() {}
}
