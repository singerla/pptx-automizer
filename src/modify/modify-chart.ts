import {
  ChartData,
  ChartColumn,
  ChartSlot,
  ChartCategory,
  ChartPoint,
  ChartBubble,
  ChartValue,
  ChartDataMapper,
} from '../types/chart-types';
import { ModificationTags, Modification } from '../types/modify-types';
import { XmlHelper } from '../helper/xml-helper';
import CellIdHelper from '../helper/cell-id-helper';
import { Workbook } from '../types/types';
import ModifyXmlHelper from '../helper/modify-xml-helper';

export class ModifyChart {
  data: ChartData;
  height: number;
  width: number;
  columns: ChartColumn[];

  sharedStrings: Document;

  workbook: ModifyXmlHelper;
  workbookTable: ModifyXmlHelper;
  chart: ModifyXmlHelper;

  constructor(
    chart: XMLDocument,
    workbook: Workbook,
    data: ChartData,
    slot: ChartSlot[],
  ) {
    this.data = data;

    this.chart = new ModifyXmlHelper(chart);
    this.workbook = new ModifyXmlHelper(workbook.sheet);
    this.workbookTable = new ModifyXmlHelper(workbook.table);

    this.sharedStrings = workbook.sharedStrings;

    this.columns = this.setColumns(slot);
    this.height = this.data.categories.length;
    this.width = this.columns.length;
  }

  modify(): void {
    this.setValues();
    this.setSeries();
    this.sliceChartSpace();

    this.prepareWorkbook();
    this.setWorkbook();
    this.sliceWorkbook();
    this.setWorkbookTable();
    this.sliceWorkbookTable();
  }

  setColumns(slot: ChartSlot[]): ChartColumn[] {
    const columns = [] as ChartColumn[];

    slot.forEach((slot) => {
      const series = slot.series;
      const index = slot.index;
      const targetCol = slot.targetCol;

      const label = slot.label ? slot.label : series.label;

      const mapData =
        slot.mapData !== undefined ? slot.mapData : (point: number) => point;

      const isStrRef = slot.isStrRef !== undefined ? slot.isStrRef : true;

      const worksheetCb = (
        point: number,
        r: number,
        category: ChartCategory,
      ): void => {
        return this.workbook.modify(
          this.rowValues(r, targetCol, mapData(point, category)),
        );
      };

      const chartCb =
        slot.type !== undefined &&
        this[slot.type] !== undefined &&
        typeof this[slot.type] === 'function'
          ? (
              point: number | ChartPoint | ChartBubble | ChartValue,
              r: number,
              category: ChartCategory,
            ): ModificationTags => {
              return this[slot.type](
                r,
                targetCol,
                point,
                category,
                slot.tag,
                mapData,
              );
            }
          : null;

      const column = <ChartColumn>{
        series: index,
        label: label,
        worksheet: worksheetCb,
        chart: chartCb,
        isStrRef: isStrRef,
      };

      columns.push(column);
    });

    return columns;
  }

  setValues(): void {
    this.data.categories.forEach((category, c) => {
      this.columns
        .filter((col) => col.chart)
        .forEach((col) => {
          if (category.values[col.series] === undefined) {
            throw new Error(
              `No value for category "${category.label}" at series "${col.label}".`,
            );
          }

          col.modTags = col.chart(category.values[col.series], c, category);

          this.chart.modify(this.series(col.series, col.modTags));
        });
    });
  }

  setSeries(): void {
    this.columns.forEach((column, colId) => {
      if (column.isStrRef === true) {
        this.chart.modify(
          this.series(column.series, {
            ...this.seriesId(column.series),
            ...this.seriesLabel(column.label, colId),
          }),
        );
      }
    });
  }

  sliceChartSpace(): void {
    this.chart.modify({
      'c:plotArea': this.slice('c:ser', this.data.series.length),
    });

    this.columns
      .filter((column) => column.modTags)
      .forEach((column) => {
        const sliceMod = {};

        Object.keys(column.modTags).forEach((tag) => {
          sliceMod[tag] = this.slice('c:pt', this.height);
        });
        this.chart.modify(this.series(column.series, sliceMod));
      });
  }

  /*
    There might be rows in an excel workbook that appear to be empty, but
    contain either no cells or none with a "v"-tag. These rows are removed
    by prepareWorkbook(). See https://github.com/singerla/pptx-automizer/issues/11
   */
  prepareWorkbook(): void {
    const rows = this.workbook.root.getElementsByTagName('row')
    for(const r in rows) {
      if(!rows[r].getElementsByTagName) continue

      const values = rows[r].getElementsByTagName('v')
      if(values.length === 0) {
        const toRemove = rows[r];
        toRemove.parentNode.removeChild(toRemove);
      }
    }
  }

  setWorkbook(): void {
    this.workbook.modify(this.spanString());
    this.workbook.modify(this.rowAttributes(0, 1));

    this.data.categories.forEach((category, c) => {
      const r = c + 1;
      this.workbook.modify(this.rowLabels(r, category.label));
      this.workbook.modify(this.rowAttributes(r, r + 1));

      this.columns.forEach((addCol) =>
        addCol.worksheet(category.values[addCol.series], r, category),
      );
    });

    this.columns.forEach((addCol, s) => {
      this.workbook.modify(this.colLabel(s + 1, addCol.label));
    });
  }

  sliceWorkbook(): void {
    this.data.categories.forEach((category, c) => {
      const r = c + 1;
      this.workbook.modify({
        row: {
          index: r,
          ...this.slice('c', this.width + 1),
        },
      });
    });

    this.workbook.modify({
      row: {
        ...this.slice('c', this.width + 1),
      },
    });

    this.workbook.modify({
      sheetData: this.slice('row', this.height + 1),
    });
  }

  series = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'c:ser': {
        index: index,
        children: children,
      },
    };
  };

  seriesId = (series: number): ModificationTags => {
    return {
      'c:idx': {
        modify: ModifyXmlHelper.attribute('val', series),
      },
      'c:order': {
        modify: ModifyXmlHelper.attribute('val', series + 1),
      },
    };
  };

  seriesLabel = (label: string, series: number): ModificationTags => {
    return {
      'c:f': {
        modify: ModifyXmlHelper.range(series + 1),
      },
      'c:v': {
        modify: ModifyXmlHelper.text(label),
      },
    };
  };

  defaultSeries(
    r: number,
    targetCol: number,
    point: number,
    category: ChartCategory,
  ): ModificationTags {
    return {
      'c:val': this.point(r, targetCol, point),
      'c:cat': this.point(r, 0, category.label),
    };
  }

  xySeries(
    r: number,
    targetCol: number,
    point: number,
    category: ChartCategory,
  ): ModificationTags {
    return {
      'c:xVal': this.point(r, targetCol, point),
      'c:yVal': this.point(r, 1, category.y),
    };
  }

  customSeries(
    r: number,
    targetCol: number,
    point: number | ChartPoint | ChartBubble | ChartValue,
    category: ChartCategory,
    tag: string,
    mapData: ChartDataMapper,
  ): ModificationTags {
    return {
      [tag]: this.point(r, targetCol, mapData(point, category)),
    };
  }

  point = (r: number, c: number, value: number | string): Modification => {
    return {
      children: {
        'c:pt': {
          index: r,
          modify: ModifyXmlHelper.value(value, r),
        },
        'c:f': {
          modify: ModifyXmlHelper.range(c, this.height),
        },
        'c:ptCount': {
          modify: ModifyXmlHelper.attribute('val', this.height),
        },
      },
    };
  };

  colLabel(c: number, label: string): ModificationTags {
    return {
      row: {
        modify: ModifyXmlHelper.attribute('spans', `1:${this.width}`),
        children: {
          c: {
            index: c,
            modify: ModifyXmlHelper.attribute(
              'r',
              CellIdHelper.getCellAddressString(c, 0),
            ),
            children: this.sharedString(label),
          },
        },
      },
    };
  }

  rowAttributes(r: number, rowId: number): ModificationTags {
    return {
      row: {
        index: r,
        modify: [
          ModifyXmlHelper.attribute('spans', `1:${this.width}`),
          ModifyXmlHelper.attribute('r', String(rowId)),
        ],
      },
    };
  }

  rowLabels(r: number, label: string): ModificationTags {
    return {
      row: {
        index: r,
        children: {
          c: {
            modify: ModifyXmlHelper.attribute(
              'r',
              CellIdHelper.getCellAddressString(0, r),
            ),
            children: this.sharedString(label),
          },
        },
      },
    };
  }

  rowValues(r: number, c: number, value: number): ModificationTags {
    return {
      row: {
        index: r,
        children: {
          c: {
            index: c,
            modify: ModifyXmlHelper.attribute(
              'r',
              CellIdHelper.getCellAddressString(c, r),
            ),
            children: this.cellValue(value),
          },
        },
      },
    };
  }

  slice(tag: string, length: number): Modification {
    return {
      children: {
        [tag]: {
          collection: (collection: HTMLCollectionOf<Element>) => {
            XmlHelper.sliceCollection(collection, length);
          },
        },
      },
    };
  }

  spanString(): ModificationTags {
    return {
      dimension: {
        modify: ModifyXmlHelper.attribute(
          'ref',
          CellIdHelper.getSpanString(0, 1, this.width, this.height),
        ),
      },
    };
  }

  cellValue(value: number): ModificationTags {
    return {
      v: {
        modify: ModifyXmlHelper.text(String(value)),
      },
    };
  }

  sharedString(label: string): ModificationTags {
    return this.cellValue(
      XmlHelper.appendSharedString(this.sharedStrings, label),
    );
  }

  setWorkbookTable(): void {
    this.workbookTable.modify({
      table: {
        modify: ModifyXmlHelper.attribute(
          'ref',
          CellIdHelper.getSpanString(0, 1, this.width, this.height),
        ),
      },
      tableColumns: {
        modify: ModifyXmlHelper.attribute('count', this.width + 1),
      },
    });

    this.columns.forEach((addCol, s) => {
      this.setWorkbookTableColumn(s + 1, addCol.label);
    });
  }

  setWorkbookTableColumn(c: number, label: string): void {
    this.workbookTable.modify({
      tableColumn: {
        index: c,
        modify: [
          ModifyXmlHelper.attribute('id', c + 1),
          ModifyXmlHelper.attribute('name', label),
        ],
      },
    });
  }

  sliceWorkbookTable(): void {
    this.workbookTable.modify({
      table: this.slice('tableColumn', this.width + 1),
    });
  }
}
