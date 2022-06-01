import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { UtilService } from '../common/util.service';

// Download Service Calss
@Injectable({
  providedIn: 'root'
})

export class DownloadexcelService 
{
  //Excel Sheet Properties
  worksheetnm: string = "";
  filenm: string = "";
  //Excel Title Properties
  showtitleflg: boolean = false;
  showtitlestyleflg: boolean = false;
  title: string = "";
  titlestyle: any = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
  //Excel Header Properties
  showheaderflg: boolean = true;
  showheaderestyleflg: boolean = false;
  headerdata: any[] = [];
  headerstyle: any = { 
                        type: 'pattern', 
                        pattern: 'solid', 
                        fgColor: { argb: 'FFE4E6EB' }
                      };
  //Excel Data Items properties                    
  lineitems: any[] = [];
  lineItemDto:any;
  // Utility Depndency Injection                    
  constructor(private utilsvc:UtilService) { }
  // Export to Excel function
  async ExportToExcel() 
  {
    //New Excel Work Book Creation
    let workbook = new Workbook();

    //add name to sheet
    let worksheet = workbook.addWorksheet(this.worksheetnm);

    // Excel Title Row Handling
    if (this.showtitleflg) 
    {
      let titleRow = worksheet.addRow([this.title]);
      if (this.showtitlestyleflg) {
        // Set font, size and style in title row.
        titleRow.font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
      }
      // Blank Row
      worksheet.addRow([]);
    }

    // Excel Header Row Handling
    if (this.showheaderflg) 
    {
      // Add Header Row
      let headerRow = worksheet.addRow(this.headerdata);

      if (this.showheaderestyleflg) 
      {
        // Cell Style : Fill and Border
        headerRow.eachCell(
          (cell, number) => {
            cell.fill = this.headerstyle;
            cell.font = {size: 12,  bold: true };
          }
        );
      }
    }

    // Converting the Data DTO in DATA array of array
    //this.lineitems= this.utilsvc.convertDtoinArray(this.lineItemDto);

    // Add Data and Conditional Formatting
    var objOutPut=[];

    for(var i=0;i<this.lineItemDto.length;i++)
    {
      const ele=this.lineItemDto[i];
      let row = worksheet.addRow(Object.values(ele));
    }
    
    // Add Data and Conditional Formatting
   /* this.lineitems.forEach
    (
      d => {
              let row = worksheet.addRow(d);
              //let qty = row.getCell(5);
              //let color = 'FF99FF99';
              //if (qty.value != undefined)
              //{
              //if (qty.value  < 500) {color = 'FF9999'}
              //}

              //qty.fill = {
              //   type: 'pattern',pattern: 'solid',
              //   fgColor: { argb: color } 
              // }
            }
    );
    */

    //add data and file name and download
    workbook.xlsx.writeBuffer()
    .then(
            (data) => 
            {
              let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
              fs.saveAs(blob, this.filenm + '_' + new Date().valueOf() + '.xlsx');
            }
          );
  }
}
