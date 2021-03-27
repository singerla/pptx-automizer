import {
  ChartData,
  ChartColumn,
  ModificationPattern,
  ModificationPatternChildren,
} from '../types/chart-types';
import { ModifyChart } from './chart';

export class ModifyChartspace extends ModifyChart {
  constructor(chart: XMLDocument, data: ChartData, columns: ChartColumn[]) {
    super(chart, data, columns);
  }

  setChart(): void {
    this.setValues();
    this.setSeries();
  }

  setValues(): void {
    this.data.categories.forEach((category, c) => {
      this.columns
        .filter((col) => col.chart)
        .forEach((col) => {
          this.pattern(
            this.series(
              col.series,
              col.chart(this, category.values[col.series], c, category),
            ),
          );
        });
    });
  }

  setSeries(): void {
    this.columns.forEach((column, colId) => {
      if (column.isStrRef) {
        this.pattern(
          this.series(column.series, {
            ...this.seriesId(column.series),
            ...this.seriesLabel(column.label, colId),
          }),
        );
      }
    });
  }

  series = (
    index: number,
    children: ModificationPatternChildren,
  ): ModificationPatternChildren => {
    return {
      'c:ser': {
        index: index,
        children: children,
      },
    };
  };

  seriesId = (series: number): ModificationPatternChildren => {
    return {
      'c:idx': {
        modify: this.attribute('val', series),
      },
      'c:order': {
        modify: this.attribute('val', series + 1),
      },
    };
  };

  seriesLabel = (
    label: string,
    series: number,
  ): ModificationPatternChildren => {
    return {
      'c:f': {
        modify: this.range(series + 1),
      },
      'c:v': {
        modify: this.text(label),
      },
    };
  };

  defaultValues = (
    label: string,
    value: number,
    index: number,
    series: number,
  ): ModificationPatternChildren => {
    return {
      'c:val': this.point(value, index, series + 1),
      'c:cat': this.point(index, 0, label),
    };
  };

  point = (
    r: number,
    c: number,
    value: number | string,
  ): ModificationPattern => {
    return {
      children: {
        'c:pt': {
          index: r,
          modify: this.value(value, r),
        },
        'c:f': {
          modify: this.range(c, this.height),
        },
        'c:ptCount': {
          modify: this.attribute('val', this.height),
        },
      },
    };
  };
}
