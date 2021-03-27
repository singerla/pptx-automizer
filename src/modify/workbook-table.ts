import { ChartData, ChartColumn } from '../types/chart-types';
import { Workbook } from '../types/types';
import { ModifyChart } from './chart';

export class ModifyWorkbookTable extends ModifyChart {
  sharedStrings: XMLDocument;
  table: Workbook['table'];

  constructor(workbook: Workbook, data: ChartData, columns?: ChartColumn[]) {
    super(workbook.table, data, columns);
    this.sharedStrings = workbook.sharedStrings;
    this.table = workbook.table;
  }

  setWorkbookTable(): void {
    this.table
      .getElementsByTagName('table')[0]
      .setAttribute('ref', this.getSpanString(0, 1, this.width, this.height));

    const tableColumns = this.table.getElementsByTagName('tableColumns')[0];
    tableColumns.setAttribute('count', String(this.width + 1));
  }

  setWorkbookTableColumn(c: number, label: string): void {
    const tableColumns = this.table.getElementsByTagName('tableColumns')[0];
    this.assertNode(tableColumns.getElementsByTagName('tableColumn'), c);

    const tableColumn = tableColumns.getElementsByTagName('tableColumn')[c];
    tableColumn.setAttribute('id', String(c + 1));
    tableColumn.setAttribute('name', label);
  }
}
