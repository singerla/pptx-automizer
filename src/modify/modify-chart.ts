import {
  ChartBubble,
  ChartCategory,
  ChartColumn,
  ChartData,
  ChartDataMapper,
  ChartPoint,
  ChartSeries,
  ChartSlot,
  ChartValueStyle,
} from '../types/chart-types';
import {
  Modification,
  ModificationTags,
  ModifyCallback,
} from '../types/modify-types';
import { XmlHelper } from '../helper/xml-helper';
import CellIdHelper from '../helper/cell-id-helper';
import { Workbook } from '../types/types';
import ModifyXmlHelper from '../helper/modify-xml-helper';
import ModifyTextHelper from '../helper/modify-text-helper';
import ModifyColorHelper from '../helper/modify-color-helper';
import { XmlDocument } from '../types/xml-types';
import { modify } from '../index';
import ModifyChartHelper from '../helper/modify-chart-helper';
import { vd } from '../helper/general-helper';

export class ModifyChart {
  data: ChartData;
  height: number;
  width: number;
  columns: ChartColumn[];

  sharedStrings: XmlDocument;

  workbook: ModifyXmlHelper;
  workbookTable: ModifyXmlHelper;
  chart: ModifyXmlHelper;

  constructor(
    chart: XmlDocument,
    workbook: Workbook,
    data: ChartData,
    slot: ChartSlot[],
  ) {
    this.data = data;

    // XmlHelper.dump(chart)

    this.chart = new ModifyXmlHelper(chart);
    this.workbook = new ModifyXmlHelper(workbook.sheet);
    this.workbookTable = workbook.table
      ? new ModifyXmlHelper(workbook.table)
      : null;

    this.sharedStrings = workbook.sharedStrings;

    this.columns = this.setColumns(slot);
    this.height = this.data.categories.length;
    this.width = this.columns.length;
  }

  modify(): void {
    this.setValues();
    this.setSeries();
    this.setSeriesDataLabels();
    this.setPointStyles();
    this.sliceChartSpace();
    this.modifyWorkbook();

    // XmlHelper.dump(this.chart.root as XmlDocument)
  }

  modifyExtended(): void {
    this.setExtData();
    this.setExtSeries();
    this.sliceExtChartSpace();
    this.modifyWorkbook();
  }

  modifyWorkbook(): void {
    this.prepareWorkbook();
    this.setWorkbook();
    this.sliceWorkbook();

    if (this.workbookTable) {
      this.setWorkbookTable();
      this.sliceWorkbookTable();
    }
  }

  setColumns(slots: ChartSlot[]): ChartColumn[] {
    const columns = [] as ChartColumn[];

    slots.forEach((slot) => {
      const series = slot.series;
      const index = slot.index;
      const targetCol = slot.targetCol;
      const targetYCol = slot.targetYCol || 1;

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
              point: number | null | ChartPoint | ChartBubble,
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
                targetYCol,
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
    this.setValuesByCategory((col) => {
      return this.series(col.series, col.modTags);
    });
  }

  setExtData(): void {
    this.setValuesByCategory((col) => {
      return {
        'cx:data': {
          children: col.modTags,
        },
      };
    });
  }

  setValuesByCategory(cb): void {
    this.data.categories.forEach((category, c) => {
      this.columns
        .filter((col) => col.chart)
        .forEach((col, s) => {
          if (category.values[col.series] === undefined) {
            throw new Error(
              `No value for category "${category.label}" at series "${col.label}".`,
            );
          }

          col.modTags = col.chart(category.values[col.series], c, category);

          this.chart.modify(cb(col));
        });
    });
  }

  setPointStyles(): void {
    const count = {};
    this.data.categories.forEach((category, c) => {
      if (category.styles) {
        category.styles.forEach((style, s) => {
          if (style === null || !Object.values(style).length) return;
          count[s] = !count[s] ? 0 : count[s];
          this.chart.modify(
            this.series(s, this.chartPoint(count[s], c, style)),
          );
          if (style.label) {
            this.chart.modify(
              this.series(s, this.chartPointLabel(count[s], c, style.label)),
            );
          }
          count[s]++;
        });
      }
    });
  }

  setSeries(): void {
    this.columns.forEach((column, colId) => {
      if (column.isStrRef === true) {
        this.chart.modify(
          this.series(column.series, {
            ...this.seriesId(column.series),
            ...this.seriesLabel(column.label, colId),
            ...this.seriesStyle(this.data.series[column.series]),
          }),
        );
      }
    });
  }

  setExtSeries(): void {
    this.columns.forEach((column, colId) => {
      if (column.isStrRef === true) {
        this.chart.modify(
          this.extSeries(column.series, {
            ...this.extSeriesLabel(column.label, colId),
          }),
        );
      }
    });
  }

  setSeriesDataLabels = (): void => {
    this.data.series.forEach((series, s) => {
      this.chart.modify(
        this.series(s, this.seriesDataLabel(s, series.style?.label)),
      );

      if (series.style?.label) {
        // Apply style for all label props helper if required
        modify.setDataLabelAttributes({
          applyToSeries: s,
          ...series.style?.label,
        })(null, this.chart.root as XmlDocument);
      }

      this.data.categories.forEach((category, c) => {
        this.chart.modify(
          this.series(s, this.seriesDataLabelsRange(c, category.label)),
        );
      });
    });
  };

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

  sliceExtChartSpace(): void {
    this.chart.modify({
      'cx:plotArea': this.slice('cx:series', this.data.series.length),
    });

    this.columns
      .filter((column) => column.modTags)
      .forEach((column) => {
        const sliceMod = {};

        Object.keys(column.modTags).forEach((tag) => {
          sliceMod[tag] = this.slice('cx:pt', this.height);
        });

        this.chart.modify({
          'cx:data': { index: column.series, children: sliceMod },
        });
      });
  }

  /*
    There might be rows in an excel workbook that appear to be empty, but
    contain either no cells or none with a "v"-tag. These rows are removed
    by prepareWorkbook(). See https://github.com/singerla/pptx-automizer/issues/11
   */
  prepareWorkbook(): void {
    const rows = this.workbook.root.getElementsByTagName('row');
    for (const r in rows) {
      if (!rows[r].getElementsByTagName) continue;

      const values = rows[r].getElementsByTagName('v');
      if (values.length === 0) {
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

  chartPoint = (
    index: number,
    idx: number,
    style: ChartValueStyle,
  ): ModificationTags => {
    if (!style?.color && !style?.border && !style?.marker) return;
    return {
      'c:dPt': {
        index: index,
        children: {
          'c:idx': {
            modify: ModifyXmlHelper.attribute('val', idx),
          },
          ...this.chartPointFill(style?.color),
          ...this.chartPointBorder(style?.border),
          ...this.chartPointMarker(style?.marker),
        },
      },
    };
  };

  chartPointFill = (color: ChartValueStyle['color']): ModificationTags => {
    if (!color?.type) return;

    return {
      'c:spPr': {
        modify: ModifyColorHelper.solidFill(color),
      },
    };
  };

  chartPointMarker = (
    markerStyle: ChartValueStyle['marker'],
  ): ModificationTags => {
    if (!markerStyle) return;

    return {
      'c:marker': {
        isRequired: false,
        children: {
          'c:spPr': {
            modify: ModifyColorHelper.solidFill(markerStyle.color),
          },
        },
      },
    };
  };
  chartPointBorder = (style: ChartValueStyle['border']): ModificationTags => {
    if (!style) return;
    const modify = <ModifyCallback[]>[];

    if (style.color) {
      modify.push(ModifyColorHelper.solidFill(style.color));
      modify.push(ModifyColorHelper.removeNoFill());
    }
    if (style.weight) {
      modify.push(ModifyXmlHelper.attribute('w', style.weight));
    }

    return {
      'a:ln': {
        modify: modify,
      },
    };
  };

  chartPointLabel = (
    index: number,
    idx: number,
    labelStyle: ChartValueStyle['label'],
  ): ModificationTags => {
    if (!labelStyle) return;

    return {
      'c:dLbls': {
        children: {
          'c:dLbl': {
            index: index,
            fromIndex: 0,
            children: {
              'c:idx': {
                modify: ModifyXmlHelper.attribute('val', String(idx)),
              },
              'a:pPr': {
                modify: ModifyColorHelper.solidFill(labelStyle?.color),
                children: {
                  'a:defRPr': {
                    isRequired: false,
                    modify: ModifyTextHelper.style(labelStyle),
                  },
                },
              },
              'a:fld': {
                children: {
                  'a:rPr': {
                    modify: [
                      ModifyColorHelper.solidFill(labelStyle?.color),
                      ModifyTextHelper.style(labelStyle),
                    ],
                  },
                  'a:defRPr': {
                    isRequired: false,
                    modify: [
                      ModifyColorHelper.solidFill(labelStyle?.color),
                      ModifyTextHelper.style(labelStyle),
                    ],
                  },
                },
              },
            },
          },
        },
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
        modify: ModifyTextHelper.content(label),
      },
    };
  };

  extSeriesLabel = (label: string, series: number): ModificationTags => {
    return {
      'cx:f': {
        modify: ModifyXmlHelper.range(series + 1),
      },
      'cx:v': {
        modify: ModifyTextHelper.content(label),
      },
    };
  };

  seriesStyle = (series: ChartSeries): ModificationTags => {
    if (!series?.style) return;

    return {
      'c:spPr': {
        modify: ModifyColorHelper.solidFill(series.style.color),
      },
      'c:marker': {
        isRequired: false,
        children: {
          'c:spPr': {
            isRequired: false,
            modify: ModifyColorHelper.solidFill(series.style.marker?.color),
          },
        },
      },
    };
  };

  seriesDataLabelsRange = (
    r: number,
    value: string | number,
  ): ModificationTags => {
    return {
      'c15:datalabelsRange': {
        isRequired: false,
        children: {
          'c:pt': {
            index: r,
            modify: ModifyXmlHelper.value(value, r),
          },
          'c15:f': {
            modify: ModifyXmlHelper.range(0, this.height),
          },
          'c:ptCount': {
            modify: ModifyXmlHelper.attribute('val', this.height),
          },
        },
      },
    };
  };

  seriesDataLabel = (s, style: ChartValueStyle['label']): ModificationTags => {
    return {
      'c:dLbls': {
        isRequired: false,
        children: {
          'a:pPr': {
            modify: ModifyColorHelper.solidFill(style?.color),
            children: {
              'a:defRPr': {
                modify: ModifyTextHelper.style(style),
              },
            },
          },
        },
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
    tag: string,
    mapData: ChartDataMapper,
    targetYCol: number,
  ): ModificationTags {
    return {
      'c:xVal': this.point(r, targetCol, point),
      'c:yVal': this.point(r, targetYCol, category.y),
    };
  }

  customSeries(
    r: number,
    targetCol: number,
    point: number | ChartPoint | ChartBubble,
    category: ChartCategory,
    tag: string,
    mapData: ChartDataMapper,
  ): ModificationTags {
    return {
      [tag]: this.point(r, targetCol, mapData(point, category)),
    };
  }

  extendedSeries(
    r: number,
    targetCol: number,
    point: number,
    category: ChartCategory,
  ): ModificationTags {
    return {
      'cx:strDim': this.extPoint(r, 0, category.label),
      'cx:numDim': this.extPoint(r, targetCol, point),
    };
  }

  extPoint = (r: number, c: number, value: string | number): Modification => {
    return {
      children: {
        'cx:pt': {
          index: r,
          modify: [
            ModifyXmlHelper.attribute('idx', r),
            ModifyXmlHelper.textContent(value),
          ],
        },
        'cx:f': {
          modify: ModifyXmlHelper.range(c, this.height),
        },
        'cx:lvl': {
          modify: ModifyXmlHelper.attribute('ptCount', this.height),
        },
      },
    };
  };

  extSeries = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'cx:series': {
        index: index,
        children: children,
      },
    };
  };

  point = (r: number, c: number, value: string | number): Modification => {
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
        fromPrevious: true,
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
        fromPrevious: true,
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
        fromPrevious: true,
        children: {
          c: {
            index: c,
            fromPrevious: true,
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
        modify: ModifyTextHelper.content(value),
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

    this.setWorkbookTableFirstColumn();
    this.columns.forEach((addCol, s) => {
      this.setWorkbookTableColumn(s + 1, addCol.label);
    });
  }

  setWorkbookTableFirstColumn(): void {
    this.workbookTable.modify({
      tableColumn: {
        index: 0,
        modify: ModifyXmlHelper.attribute('id', 1),
      },
    });
  }

  setWorkbookTableColumn(c: number, label: string): void {
    this.workbookTable.modify({
      tableColumn: {
        index: c,
        fromPrevious: true,
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
