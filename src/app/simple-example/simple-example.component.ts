import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-simple-example',
  template: `
  <div class="row mt-3">

    <div class="col">
      <button class="btn btn-primary mr-2" (click)="exportar()">
        Exportar Tabela
      </button>
    </div>

  </div>

  <div class="row mt-3">

    <div class="col">

      <div class="table-responsive">

        <table class="table table-bordered table-sm table-hover table-dark table-striped" #exemploTabela>
          <thead>
            <tr>
              <th>Coluna A</th>
              <th>Coluna B</th>
              <th>Coluna C</th>
              <th>Coluna D</th>
            </tr>
          </thead>

          <tbody>
            <tr>
              <td>Linha 1</td>
              <td>valor 1</td>
              <td>valor 2</td>
              <td>valor 3</td>
            </tr>
            <tr>
              <td>Linha 2</td>
              <td>valor 1</td>
              <td>valor 2</td>
              <td>valor 3</td>
            </tr>
            <tr>
              <td>Linha 3</td>
              <td>valor 1</td>
              <td>valor 2</td>
              <td>valor 3</td>
            </tr>
          </tbody>
        </table>

      </div>

    </div>

  </div>
  `,
  styles: ['']
})
export class SimpleExampleComponent implements OnInit {

  @ViewChild('exemploTabela', { static: false }) exemploTabela: ElementRef;

  constructor() { }

  ngOnInit(): void {
  }

  exportar() {
    let fileName: string = 'exemplo_tabela.xlsx';
    let copyTable = this.exemploTabela.nativeElement.cloneNode(true);

    const worksheet: XLSX.WorkSheet = XLSX.utils.table_to_sheet(copyTable);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    XLSX.writeFile(workbook, fileName);
  }

}
