import { ModifyChart } from '../modify/modify-chart';
import { Workbook } from '../types/types';
import {
  ChartData,
  ChartBubble,
  ChartSlot,
  ChartCategory,
  ChartSeries,
} from '../types/chart-types';

export default class ModifyChartHelper {
  /**
   * Set chart data to modify default chart types.
   * See `__tests__/modify-existing-chart.test.js`
   */
  static setChartData = (data: ChartData) => (
    element: XMLDocument,
    chart: Document,
    workbook: Workbook,
  ): void => {
    const slots = [] as ChartSlot[];
    data.series.forEach((series: ChartSeries, s: number) => {
      slots.push({
        index: s,
        series: series,
        targetCol: s + 1,
        type: 'defaultSeries',
      });
    });

    new ModifyChart(chart, workbook, data, slots).modify();

    // XmlHelper.dump(chart)
    // XmlHelper.dump(workbook.table)
  };

  /**
   * Set chart data to modify vertical line charts.
   * See `__tests__/modify-chart-vertical-lines.test.js`
   */
  static setChartVerticalLines = (data: ChartData) => (
    element: XMLDocument,
    chart: Document,
    workbook: Workbook,
  ): void => {
    const slots = [] as ChartSlot[];

    slots.push({
      label: `Y-Values`,
      mapData: (point: number, category: ChartCategory) => category.y,
      targetCol: 1,
    });

    data.series.forEach((series: ChartSeries, s: number) => {
      slots.push({
        index: s,
        series: series,
        targetCol: s + 2,
        type: 'xySeries',
      });
    });

    new ModifyChart(chart, workbook, data, slots).modify();
  };

  /**
   * Set chart data to modify bubble charts.
   * See `__tests__/modify-chart-bubbles.test.js`
   */
  static setChartBubbles = (data: ChartData) => (
    element: XMLDocument,
    chart: Document,
    workbook: Workbook,
  ): void => {
    const slots = [] as ChartSlot[];

    data.series.forEach((series: ChartSeries, s: number) => {
      const colId = s * 3;
      slots.push({
        index: s,
        series: series,
        targetCol: colId + 1,
        type: 'customSeries',
        tag: 'c:xVal',
        mapData: (point: ChartBubble): number => point.x,
      });
      slots.push({
        label: `${series.label}-Y-Value`,
        index: s,
        series: series,
        targetCol: colId + 2,
        type: 'customSeries',
        tag: 'c:yVal',
        mapData: (point: ChartBubble): number => point.y,
        isStrRef: false,
      });
      slots.push({
        label: `${series.label}-Size`,
        index: s,
        series: series,
        targetCol: colId + 3,
        type: 'customSeries',
        tag: 'c:bubbleSize',
        mapData: (point: ChartBubble): number => point.size,
        isStrRef: false,
      });
    });

    new ModifyChart(chart, workbook, data, slots).modify();

    // XmlHelper.dump(chart)
  };
}
