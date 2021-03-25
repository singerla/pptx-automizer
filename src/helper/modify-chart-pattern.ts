import {
  ChartData,
  ModificationPattern,
  ModificationPatternChildren,
  XYChartData,
} from '../types/types';
import { ModifyChart } from './modify-chart';

export class ModifyChartPattern extends ModifyChart {
  valuesPattern: any;

  constructor(chart: XMLDocument, data: XYChartData | ChartData, valuesPattern, addCols?: any[]) {
    super(chart, data, addCols)
    this.valuesPattern = valuesPattern
  }

  setChart() {
    this.setValues()
    this.setSeries()
  }

  setValues() {
    this.data.categories.forEach((category, c) => {
      category.values.forEach((value, s) => {
        this.pattern(this.series(s, this.valuesPattern(this, category, value, c, s)));
      });
    });
  }

  setSeries() {
    this.data.series.forEach((series, s) => {
      this.pattern(this.series(s, { 
        ...this.seriesId(s),
        ...this.seriesLabel(series.label, s + this.addColsLength)
      }));
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
        modify: this.attribute('val', series+1),
      }
    };
  };

  seriesLabel = (label: string, series: number): ModificationPatternChildren => {
    return {
      'c:f': {
        modify: this.range(series+1),
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
      'c:val': this.point(value, index, series+1),
      'c:cat': this.point(label, index, 0),
    };
  };

  xyValues = (
    xValue: number,
    yValue: number,
    index: number,
    series: number,
  ): ModificationPatternChildren => {
    return {
      'c:xVal': this.point(xValue, index, series+2),
      'c:yVal': this.point(yValue, index, 1),
    };
  };

  point = (
    value: number | string,
    index: number,
    colId: number,
  ): ModificationPattern => {
    return {
      children: {
        'c:pt': {
          index: index,
          modify: this.value(value, index),
        },
        'c:f': {
          modify: this.range(colId, this.height),
        },
        'c:ptCount': {
          modify: this.attribute('val', this.height),
        },
      },
    };
  };
}
