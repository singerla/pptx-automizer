import { XmlHelper } from './xml-helper';
import { ChartData, ChartColumn, ChartBubble } from '../types/chart-types';
import { FrameCoordinates, Workbook } from '../types/types';
import { ModifyChartspace } from '../modify/chartspace';
import { ModifyWorkbook } from '../modify/workbook';

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
  const columns = [] as ChartColumn[];

  data.series.forEach((series, s) => {
    columns.push({
      series: s,
      label: `${series.label}`,
      worksheet: (ctx, point: number, r: number) =>
        ctx.pattern(ctx.rowValues(r, s + 1, point)),
      chart: (ctx, point: number, r: number, category) => {
        return {
          'c:val': ctx.point(r, s + 1, point),
          'c:cat': ctx.point(r, 0, category.label),
        };
      },
      isStrRef: true,
    });
  });

  new ModifyChartspace(chart, data, columns).setChart();
  new ModifyWorkbook(workbook, data, columns).setWorkbook();
};

export const setChartVerticalLines = (data: ChartData) => (
  element: XMLDocument,
  chart: Document,
  workbook: Workbook,
): void => {
  const columns = [] as ChartColumn[];

  columns.push({
    label: `Y-Values`,
    worksheet: (ctx, point, r: number, category) =>
      ctx.pattern(ctx.rowValues(r, 1, category.y)),
  });

  data.series.forEach((series, s) => {
    columns.push({
      series: s,
      label: `${series.label}`,
      worksheet: (ctx, point: number, r: number) =>
        ctx.pattern(ctx.rowValues(r, s + 2, point)),
      chart: (ctx, point: number, r: number, category) => {
        return {
          'c:xVal': ctx.point(r, s + 2, point),
          'c:yVal': ctx.point(r, 1, category.y),
        };
      },
      isStrRef: true,
    });
  });

  new ModifyChartspace(chart, data, columns).setChart();
  new ModifyWorkbook(workbook, data, columns).setWorkbook();
};

export const setChartBubbles = (data: ChartData) => (
  element: XMLDocument,
  chart: Document,
  workbook: Workbook,
): void => {
  const columns = [] as ChartColumn[];

  data.series.forEach((series, s) => {
    const colId = s * 3;
    columns.push({
      series: s,
      label: `${series.label}`,
      worksheet: (ctx, point: ChartBubble, r: number) =>
        ctx.pattern(ctx.rowValues(r, colId + 1, point.x)),
      chart: (ctx, point: ChartBubble, r: number) => {
        return { 'c:xVal': ctx.point(r, colId + 1, point.x) };
      },
      isStrRef: true,
    });
    columns.push({
      series: s,
      label: `${series.label}-Y-Value`,
      worksheet: (ctx, point: ChartBubble, r: number) =>
        ctx.pattern(ctx.rowValues(r, colId + 2, point.y)),
      chart: (ctx, point: ChartBubble, r: number) => {
        return { 'c:yVal': ctx.point(r, colId + 2, point.y) };
      },
    });
    columns.push({
      series: s,
      label: `${series.label}-Size`,
      worksheet: (ctx, point: ChartBubble, r: number) =>
        ctx.pattern(ctx.rowValues(r, colId + 3, point.size)),
      chart: (ctx, point: ChartBubble, r: number) => {
        return { 'c:bubbleSize': ctx.point(r, colId + 3, point.size) };
      },
    });
  });

  new ModifyChartspace(chart, data, columns).setChart();
  new ModifyWorkbook(workbook, data, columns).setWorkbook();
};

export const dump = (element: XMLDocument | Document): void => {
  XmlHelper.dump(element);
};
