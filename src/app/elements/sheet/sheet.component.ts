import { Component, OnInit } from '@angular/core';
import * as Excel from 'exceljs';

export interface Tile {
  color: string;
  cols: number;
  rows: number;
  text: string;
}

@Component({
  selector: 'ss-sheet',
  templateUrl: 'sheet.component.html',
  styles: [
  ]
})
export class SheetComponent implements OnInit {

  tiles: Tile[] = [];
  colCount: number = 10;

  constructor() { }

  ngOnInit(): void {
    
  }

  readExcel(event) {
    const workbook = new Excel.Workbook();
    const target: DataTransfer = <DataTransfer>(event.target);
    
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }

    /**
     * Final Solution For Importing the Excel FILE
    */

    const arrayBuffer = new Response(target.files[0]).arrayBuffer();

    arrayBuffer.then((data) => {
      workbook.xlsx.load(data)
        .then(() => {

          // play with workbook and worksheet now
          console.log(workbook);

          const worksheet = workbook.getWorksheet(1);
          this.colCount = worksheet.columnCount;
          const rowCount = worksheet.rowCount;

          console.log(this.colCount);

          

          for(let ri = 1; ri <= rowCount; ri++) {
            const row: Excel.Row = worksheet.getRow(ri);
            for (let ci = 1; ci <= row.cellCount; ci++) {
              const cell: Excel.Cell = row.getCell(ci);
              const tile: Tile = { text: cell.text, cols: 1, rows: 1, color: 'lightblue' };
              this.tiles.push(tile);
            }
          }
        });
    });
  }
}
