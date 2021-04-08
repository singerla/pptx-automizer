import { XmlHelper } from './xml-helper';
import { ModifyChart } from '../modify/modify-chart';
import { Workbook } from '../types/types';
import {
  ChartData,
  ChartBubble,
  ChartSlot,
  ChartCategory,
  ChartSeries,
} from '../types/chart-types';
import { TableData } from '../types/table-types';
import { ShapeCoordinates } from '../types/shape-types';

export const setSolidFill = (element: XMLDocument): void => {
  element
    .getElementsByTagName('a:solidFill')[0]
    .getElementsByTagName('a:schemeClr')[0]
    .setAttribute('val', 'accent6');
};

export const setText = (text: string) => (element: XMLDocument): void => {
  element.getElementsByTagName('a:t')[0].firstChild.textContent = text;
};

// e.g. setPosition({x: 8000000, h:5000000})
export const setPosition = (pos: ShapeCoordinates) => (
  element: XMLDocument,
): void => {
  const map = {
    x: { tag: 'a:off', attribute: 'x' },
    y: { tag: 'a:off', attribute: 'y' },
    w: { tag: 'a:ext', attribute: 'cx' },
    h: { tag: 'a:ext', attribute: 'cy' },
  };

  const parent = 'a:xfrm';

  Object.keys(pos).forEach((key) => {
    element
      .getElementsByTagName(parent)[0]
      .getElementsByTagName(map[key].tag)[0]
      .setAttribute(map[key].attribute, pos[key]);
  });
};

export const setAttribute = (
  tagName: string,
  attribute: string,
  value: string | number,
  count?: number,
) => (element: XMLDocument): void => {
  element
    .getElementsByTagName(tagName)
    [count || 0].setAttribute(attribute, String(value));
};

export const setChartData = (data: ChartData) => (
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

export const setChartVerticalLines = (data: ChartData) => (
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

export const setChartBubbles = (data: ChartData) => (
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

export const setTableData = (data: TableData) => (
  element: XMLDocument | Document | Element,
): void => {
  XmlHelper.dump(element);
};

export const dump = (element: XMLDocument | Document | Element): void => {
  XmlHelper.dump(element);
};
