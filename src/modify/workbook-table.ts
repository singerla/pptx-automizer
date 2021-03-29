import CellIdHelper from '../helper/cell-id-helper';
import { ChartData, ChartColumn } from '../types/chart-types';
import { Workbook } from '../types/types';
import { ModifyChart } from './chart';

export class ModifyWorkbookTable extends ModifyChart {
  constructor(workbook: Workbook, data: ChartData, columns?: ChartColumn[]) {
    super(workbook.table, data, columns);
  }

  setWorkbookTable(): void {
    this.pattern({
      'table': {
        modify: this.attribute('ref', CellIdHelper.getSpanString(0, 1, this.width, this.height))
      },
      'tableColumns': {
        modify: this.attribute('count', this.width + 1)
      }
    })
  }

  setWorkbookTableColumn(c: number, label: string): void {
    this.pattern({
      'tableColumn': {
        index: c,
        modify: [
          this.attribute('id', c + 1),
          this.attribute('name', label),
        ]
      }
    })
  }
}
