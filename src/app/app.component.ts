import { Component, ViewChild, ElementRef, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import { DecimalPipe } from '@angular/common';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  @ViewChild('minhaTabela', {static: false}) minhaTabela: ElementRef;
  @ViewChild('minhaTabela2', {static: false}) minhaTabela2: ElementRef;

  public decimalPipe = new DecimalPipe('pt-PT');

  public dadosTabela1 = {
    "body":[
       {
          "name":"Conjunto A",
          "show":true,
          "row":2,
          "col":1,
          "metrica":"Valor 1",
          "values":[
             7.582089400770542e-14,
             2.6325932545034905e-13,
             7.303807336711543e-13
          ]
       },
       {
          "name":"Conjunto A",
          "show":false,
          "row":1,
          "col":1,
          "metrica":"Valor 2",
          "values":[
             -546.1315446214965,
             -564.3359294422132,
             -6174.867473487933
          ]
       },
       {
          "name":"Conjunto B",
          "show":true,
          "row":1,
          "col":2,
          "metrica":"",
          "values":[
            -737.1685230905573,
            -761.7408071935756,
            -8281.782871348823
        ]
      },
       {
          "name":"Conjunto C",
          "show":true,
          "row":2,
          "col":1,
          "metrica":"Valor 1",
          "values":[
            -737.1685230905573,
            -761.7408071935756,
            -8281.782871348823
         ]
       },
       {
          "name":"Conjunto C",
          "show":false,
          "row":1,
          "col":1,
          "metrica":"Valor 2",
          "values":[
            542.8896003390456,
            560.985920350347,
            6121.2911785189135
         ]
       }
    ],
    "header":[
       "01/2020",
       "02/2020",
       "03/2020",
    ]
  };

  public dadosTabela2 = {
    "body":[
       {
          "name":"Conjunto A",
          "show":true,
          "row":1,
          "col":1,
          "metrica":"",
          "values":[
             7.582089400770542e-14,
             2.6325932545034905e-13,
             7.303807336711543e-13
          ]
       },
       {
          "name":"Conjunto B",
          "show":true,
          "row":1,
          "col":1,
          "metrica":"",
          "values":[
            -737.1685230905573,
            -761.7408071935756,
            -8281.782871348823
        ]
      },
       {
          "name":"Conjunto C",
          "show":true,
          "row":1,
          "col":1,
          "metrica":"",
          "values":[
            -737.1685230905573,
            -761.7408071935756,
            -8281.782871348823
         ]
       },
    ],
    "header":[
       "01/2020",
       "02/2020",
       "03/2020",
    ]
  };
  
  constructor() {}

  ngOnInit() {

  }

  exportar(tabelas: Array<string>) {
    if (tabelas.length === 1) {
      this.exportarUmaTabela(tabelas[0]);
    } else if (tabelas.length > 1) {
      this.exportarVariasTabelas(tabelas);
    }
  }

  exportarUmaTabela(tabela: string) {
    const obj = this.getDados()[tabela];

    let fileName: string = `${obj.name}.xlsx`;
    let copyTable = obj.element.cloneNode(true);

    copyTable.deleteTFoot();

    /** Fazer alterações em alguma célula antes de exportar */

    let rowsTable = copyTable.querySelector("tbody").querySelectorAll("tr");

    rowsTable.forEach(row => {
      row.querySelectorAll('td').forEach((value, index) => {
        value.innerHTML = value.innerHTML.replace('.', '').replace(',', '.');
      });
    });

    const worksheet: XLSX.WorkSheet=XLSX.utils.table_to_sheet(copyTable);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    XLSX.writeFile(workbook, fileName);
  }

  exportarVariasTabelas(tabelas: Array<string>) {
    let fileName: string = 'Tabela_A_e_B.xlsx';

    let copyTable = document.createElement('div');

    tabelas.forEach((tabela) => {
      copyTable.appendChild(this.getDados()[tabela].element.cloneNode(true));
    });

    /** Fazer alterações em alguma célula antes de exportar */

    let tables = copyTable.querySelectorAll('table');

    tables.forEach(table => {
      table.deleteTFoot();
      table.querySelector("tbody").insertRow(); // Separar tabelas por uma linha vazia
      let rowsTable = table.querySelector("tbody").querySelectorAll("tr");

      rowsTable.forEach((row, iRow) => {
        row.querySelectorAll('td').forEach((value, iCol) => {
          value.innerHTML = value.innerHTML.replace('.', '').replace(',', '.');
        });
      });
    });

    const worksheet: XLSX.WorkSheet=XLSX.utils.table_to_sheet(copyTable);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    XLSX.writeFile(workbook, fileName);
  }

  getDados(): object {
    return {
      tabela_A: {
        name: 'Nome_Tabela_A',
        element: this.minhaTabela.nativeElement
      },
      tabela_B: {
        name: 'Nome_Tabela_B',
        element: this.minhaTabela2.nativeElement
      }
    }
  }
}
