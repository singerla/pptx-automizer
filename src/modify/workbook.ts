import {
  ChartData,
  ChartColumn,
  ModificationPatternChildren,
} from '../types/chart-types';
import { Workbook } from '../types/types';
import { ModifyChart } from './chart';
import { ModifyWorkbookTable } from './workbook-table';

export class ModifyWorkbook extends ModifyChart {
  sharedStrings: XMLDocument;
  table: ModifyWorkbookTable;

  constructor(workbook: Workbook, data: ChartData, columns?: ChartColumn[]) {
    super(workbook.sheet, data, columns);
    this.sharedStrings = workbook.sharedStrings;
    this.table = new ModifyWorkbookTable(workbook, data, columns);
  }

  setWorkbook(): void {
    this.pattern(this.spanString());

    this.data.categories.forEach((category, c) => {
      const r = c + 1;
      this.pattern(this.rowLabels(r, category.label));
      this.pattern(this.rowAttributes(r, r + 1));

      this.columns.forEach((addCol) =>
        addCol.worksheet(this, category.values[addCol.series], r, category),
      );
    });

    this.pattern(this.rowAttributes(0, 1));
    this.table.setWorkbookTable();

    this.columns.forEach((addCol, s) => {
      this.pattern(this.colLabel(s + 1, addCol.label));
      this.table.setWorkbookTableColumn(s + 1, addCol.label);
    });
  }

  colLabel(c: number, label: string): ModificationPatternChildren {
    return {
      row: {
        modify: this.attribute('spans', `1:${this.width}`),
        children: {
          c: {
            index: c,
            modify: this.attribute('r', this.getCellAddressString(c, 0)),
            children: this.sharedString(label),
          },
        },
      },
    };
  }

  rowAttributes(r: number, rowId: number): ModificationPatternChildren {
    return {
      row: {
        index: r,
        modify: [
          this.attribute('spans', `1:${this.width}`),
          this.attribute('r', String(rowId)),
        ],
      },
    };
  }

  rowLabels(r: number, label: string): ModificationPatternChildren {
    return {
      row: {
        index: r,
        children: {
          c: {
            modify: this.attribute('r', this.getCellAddressString(0, r)),
            children: this.sharedString(label),
          },
        },
      },
    };
  }

  rowValues(r: number, c: number, value: number): ModificationPatternChildren {
    return {
      row: {
        index: r,
        children: {
          c: {
            index: c,
            modify: this.attribute('r', this.getCellAddressString(c, r)),
            children: this.cellValue(value),
          },
        },
      },
    };
  }

  spanString(): ModificationPatternChildren {
    return {
      dimension: {
        modify: this.attribute(
          'ref',
          this.getSpanString(0, 1, this.width, this.height),
        ),
      },
    };
  }

  sharedString(label: string): ModificationPatternChildren {
    const stringId = this.appendSharedString(this.sharedStrings, label);
    return this.cellValue(stringId);
  }

  cellValue(value: number): ModificationPatternChildren {
    return {
      v: {
        modify: this.text(String(value)),
      },
    };
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
