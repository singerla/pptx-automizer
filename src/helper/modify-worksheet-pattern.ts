import {
  ChartData,
  ModificationPatternChildren,
  Workbook,
  XYChartData,
} from '../types/types';
import { ModifyChart } from './modify-chart';

export class ModifyWorkbookPattern extends ModifyChart {
  sharedStrings: XMLDocument;
  table: any;

  constructor(workbook: Workbook, data: XYChartData | ChartData, addCols?: any[]) {
    super(workbook.sheet, data, addCols)
    this.sharedStrings = workbook.sharedStrings
    this.table = workbook.table
  }

  setWorkbook = () => {
    this.pattern(this.spanString())
  
    this.data.categories.forEach((category, c) => {
      const r = c+1
      this.pattern(this.rowLabels(r, category.label));
      this.pattern(this.rowAttributes(r, r+1));
      
      this.addCols.forEach(addCol => addCol(this, category, r))

      category.values.forEach((xValue, s) => {
        this.pattern(this.rowValues(r, s+1+this.addColsLength, xValue));
      });
    });
  
    this.pattern(this.rowAttributes(0, 1));
    this.data.series.forEach((series, s) => {
      this.pattern(this.colLabel(s+1+this.addColsLength, series.label))
    });
  }

  colLabel = (c:number, label:string): ModificationPatternChildren => {
    this.setTableColumn(c, label)
    return {
      'row': {
        modify: this.attribute('spans', `1:${this.width}`),
        children: {
          'c': {
            index: c,
            modify: this.attribute('r', this.getCellAddressString(c,0)),
            children: this.sharedString(label)
          }
        }
      }
    }
  }

  rowAttributes = (r: number, rowId: number): ModificationPatternChildren => {
    return {
      'row': {
        index: r,
        modify: [
          this.attribute('spans', `1:${this.width}`),
          this.attribute('r', String(rowId))
        ]
      }
    }
  }

  rowLabels = (r:number, label:string): ModificationPatternChildren => {
    return {
      'row': {
        index: r,
        children: {
          'c': {
            modify: this.attribute('r', this.getCellAddressString(0,r)),
            children: this.sharedString(label)
          }
        }
      }
    }
  }

  rowValues = (r:number, c:number, value:number): ModificationPatternChildren => {
    return {
      'row': {
        index: r,
        children: {
          'c': {
            index: c,
            modify: this.attribute('r', this.getCellAddressString(c,r)),
            children: this.cellValue(value)
          }
        }
      }
    }
  }

  spanString(): ModificationPatternChildren {
    return {
      'dimension': {
        modify: this.attribute('ref', this.getSpanString(0, 1, this.width - this.addColsLength, this.height))
      }
    }
  }

  sharedString = (label: string): ModificationPatternChildren => {
    const stringId = this.appendSharedString(
      this.sharedStrings,
      label
    );
    return this.cellValue(stringId)
  };

  cellValue = (value: number): ModificationPatternChildren => {
    return {
      'v': {
        modify: this.text(String(value))
      }
    }
  };

  setTableColumn(c: number, label: string) {
    this.table.getElementsByTagName('table')[0]
      .setAttribute('ref', this.getSpanString(0, 1, this.width-1, this.height))

    const tableColumns = this.table.getElementsByTagName('tableColumns')[0]
    tableColumns.setAttribute('count', this.width-1)

    this.assert(tableColumns.getElementsByTagName('tableColumn') , c)

    const tableColumn = tableColumns.getElementsByTagName('tableColumn')[c]
    tableColumn.setAttribute('id', c+1)
    tableColumn.setAttribute('name', label);
  }

  appendSharedString(sharedStrings: Document, stringValue: string): number {
    const strings = sharedStrings.getElementsByTagName('sst')[0];
    const newLabel = sharedStrings.createTextNode(stringValue);
    const newText = sharedStrings.createElement('t');
    newText.appendChild(newLabel);

    const newString = sharedStrings.createElement('si');
    newString.appendChild(newText);

    strings.appendChild(newString);

    return strings.getElementsByTagName('si').length - 1;
  }
}
