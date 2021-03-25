import { XmlHelper } from './xml-helper';
import {
  ChartData,
  FrameCoordinates,
  Workbook,
  XYChartData,
} from '../types/types';

import { ModifyChartPattern } from './modify-chart-pattern';
import { ModifyWorkbookPattern } from './modify-worksheet-pattern';

export const setSolidFill = (element: XMLDocument): void => {
  element
    .getElementsByTagName('a:solidFill')[0]
    .getElementsByTagName('a:schemeClr')[0]
    .setAttribute('val', 'accent6');
};

export const setText = (text: string) => (element: XMLDocument): void => {
  element.getElementsByTagName('a:t')[0].firstChild.textContent = text;
};

// eslint-disable-next-line @typescript-eslint/no-unused-vars
export const revertElements = (slide: Document): void => {
  // dump(slide)
};

// e.g. setPosition({x: 8000000, h:5000000})
export const setPosition = (pos: FrameCoordinates) => (
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
  const valuesPattern = (ctx, category, value, c, s) => ctx.defaultValues(category.label, value, c, s)
  const modCh = new ModifyChartPattern(chart, data, valuesPattern);
  modCh.setChart()
  
  const modWbk = new ModifyWorkbookPattern(workbook, data);
  modWbk.setWorkbook()
};

export const setChartVerticalLines = (data: XYChartData) => (
  element: XMLDocument,
  chart: Document,
  workbook: Workbook,
): void => {
  const valuesPattern = (ctx, category, value, c, s) => ctx.xyValues(value, category.yValue, c, s)
  const addCols = [
    (ctx, category, r:number) => ctx.pattern(ctx.rowValues(r, 1, category.yValue))
  ]

  const modCh = new ModifyChartPattern(chart, data, valuesPattern, addCols);
  modCh.setChart()

  const modWbk = new ModifyWorkbookPattern(workbook, data, addCols);
  modWbk.setWorkbook()
};

export const setChartBubbles = (data: XYChartData) => (
  element: XMLDocument,
  chart: Document,
  workbook: Workbook,
): void => {
  const valuesPattern = (ctx, category, value, c, s) => ctx.xyValues(value, category.yValue, c, s)
  const addCols = [
    (ctx, category, r:number) => ctx.pattern(ctx.rowValues(r, 1, category.yValue)),
    (ctx, category, r:number) => ctx.pattern(ctx.rowValues(r, 5, category.sizes[0])),
    (ctx, category, r:number) => ctx.pattern(ctx.rowValues(r, 6, category.sizes[1])),
    (ctx, category, r:number) => ctx.pattern(ctx.rowValues(r, 7, category.sizes[2]))
  ]

  const modCh = new ModifyChartPattern(chart, data, valuesPattern, addCols);
  modCh.setChart()

  const modWbk = new ModifyWorkbookPattern(workbook, data, addCols);
  modWbk.setWorkbook()
};

export const dump = (element: XMLDocument | Document): void => {
  XmlHelper.dump(element);
};
