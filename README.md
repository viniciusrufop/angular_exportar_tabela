# Exportar tabelas html para excel com angular

Para testar funcionalidade bastar acessar o link [Teste da aplicação](https://viniciusrufop.github.io/angular_exportar_tabela_excel/)

## Instalar dependência
Rodar comando `npm install xlsx`.

## Exemplo simples
```
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

```

# Build para usar github page

Run `ng build --prod --base-href=https://viniciusrufop.github.io/angular_exportar_tabela_excel/`

# AngularExportaTabelaExcel

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 9.1.4.

## Development server

Run `ng serve` for a dev server. Navigate to `http://localhost:4200/`. The app will automatically reload if you change any of the source files.

## Code scaffolding

Run `ng generate component component-name` to generate a new component. You can also use `ng generate directive|pipe|service|class|guard|interface|enum|module`.

## Build

Run `ng build` to build the project. The build artifacts will be stored in the `dist/` directory. Use the `--prod` flag for a production build.

## Running unit tests

Run `ng test` to execute the unit tests via [Karma](https://karma-runner.github.io).

## Running end-to-end tests

Run `ng e2e` to execute the end-to-end tests via [Protractor](http://www.protractortest.org/).

## Further help

To get more help on the Angular CLI use `ng help` or go check out the [Angular CLI README](https://github.com/angular/angular-cli/blob/master/README.md).
