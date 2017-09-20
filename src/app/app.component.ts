import { Component } from '@angular/core';
import * as OfficeHelpers from '@microsoft/office-js-helpers';
import { Data } from './data';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  perf = '0';
  perf2;

  createTable() {

    Excel.run(async (context) => {

      // create new ExcelSheet
      await OfficeHelpers.ExcelUtilities.forceCreateSheet(context.workbook, 'Example1');

      const sheet = context.workbook.worksheets.getItem('Example1');
      const t0 = performance.now();
      const example1Table = sheet.tables.add('A1:Q1', true);
      example1Table.name = 'Example1';

      example1Table.getHeaderRowRange().values = [[
        'Col1', 'Col2', 'Col3', 'Col4', 'col5', 'col6', 'col7',
        'col8', 'col9', 'col10', 'Col11', 'Col12', 'Col13', 'Col14',
        'col15', 'Col16', 'Col17'
      ]];

      // Doing it this way is slow from what I have observed
      // at 2000 records 12 columns it can take 12 seconds.
      // crashes when trying to do 50k records
      example1Table.rows.add(null, this.preppedData());

      sheet.activate();
      await context.sync();
      const t1 = performance.now();
      this.perf = (t1 - t0).toFixed(0);
      console.log(`Took ${this.perf} milliseconds to generate`);

    });

  }

  createTable2() {
    Excel.run(async (context) => {

      // create new ExcelSheet
      await OfficeHelpers.ExcelUtilities.forceCreateSheet(context.workbook, 'Example2');

      const sheet = context.workbook.worksheets.getItem('Example2');
      const t0 = performance.now();

      // This time create the table so it is as large as the anticipated data
      // in this case it is 2000 records and add 1 to account for header
      const example2Table = sheet.tables.add('A1:Q2001', true);
      example2Table.name = 'Example2';

      example2Table.getHeaderRowRange().values = [[
        'Col1', 'Col2', 'Col3', 'Col4', 'col5', 'col6', 'col7',
        'col8', 'col9', 'col10', 'Col11', 'Col12', 'Col13', 'Col14',
        'col15', 'Col16', 'Col17'
      ]];

      // This time I grab the BodyRange and setting the Values of the range
      // This seems to be much faster than adding it to the RowsCollection
      // I was able to get this to scale to 100k records but it still took
      // 1 minute 30 seconds
      example2Table.getDataBodyRange().values = this.preppedData();

      sheet.activate();
      await context.sync();
      const t1 = performance.now();
      this.perf2 = (t1 - t0).toFixed(0);
      console.log(`Took ${this.perf} milliseconds to generate`);

    });
  }


  /**
   * Grabs Mocked data and converts to Multi-dimensional Array [][]
   */
  preppedData() {
    return Data.dataset1.map(row =>
      Object.keys(row).map(key => row[key])
    );
  }

}
