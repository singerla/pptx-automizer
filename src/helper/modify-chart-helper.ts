import { ModifyChart } from '../modify/modify-chart';
import { Workbook } from '../types/types';
import {
  ChartAxisRange,
  ChartBubble,
  ChartCategory,
  ChartData,
  ChartPlotArea,
  ChartPoint,
  ChartSeries,
  ChartSlot,
} from '../types/chart-types';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { disconnect } from 'process';

export default class ModifyChartHelper {
  /**
   * Set chart data to modify default chart types.
   * See `__tests__/modify-existing-chart.test.js`
   */
  static setChartData =
    (data: ChartData) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
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
  static setChartVerticalLines =
    (data: ChartData) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
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
   * Set chart data to modify scatter charts.
   * See `__tests__/modify-chart-scatter.test.js`
   */
  static setChartScatter =
    (data: ChartData) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
    ): void => {
      const slots = [] as ChartSlot[];

      data.series.forEach((series: ChartSeries, s: number) => {
        const colId = s * 2;
        slots.push({
          index: s,
          series: series,
          targetCol: colId + 1,
          type: 'customSeries',
          tag: 'c:xVal',
          mapData: (point: ChartPoint): number => point.x,
        });
        slots.push({
          label: `${series.label}-Y-Value`,
          index: s,
          series: series,
          targetCol: colId + 2,
          type: 'customSeries',
          tag: 'c:yVal',
          mapData: (point: ChartPoint): number => point.y,
          isStrRef: false,
        });
      });

      new ModifyChart(chart, workbook, data, slots).modify();

      // XmlHelper.dump(chart)
    };

  /**
   * Set chart data to modify combo charts.
   * This type is prepared for
   * first series: bar chart (e.g. total)
   * other series: vertical lines
   * See `__tests__/modify-chart-scatter.test.js`
   */
  static setChartCombo =
    (data: ChartData) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
    ): void => {
      const slots = [] as ChartSlot[];

      slots.push({
        index: 0,
        series: data.series[0],
        targetCol: 1,
        type: 'defaultSeries',
      });

      slots.push({
        index: 1,
        label: `Y-Values`,
        mapData: (point: number, category: ChartCategory) => category.y,
        targetCol: 2,
      });

      data.series.forEach((series: ChartSeries, s: number) => {
        if (s > 0)
          slots.push({
            index: s,
            series: series,
            targetCol: s + 2,
            targetYCol: 2,
            type: 'xySeries',
          });
      });

      new ModifyChart(chart, workbook, data, slots).modify();
    };

  /**
   * Set chart data to modify bubble charts.
   * See `__tests__/modify-chart-bubbles.test.js`
   */
  static setChartBubbles =
    (data: ChartData) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
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

  /**
   * Set chart data to modify extended chart types.
   * See `__tests__/modify-existing-extended-chart.test.js`
   */
  static setExtendedChartData =
    (data: ChartData) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
    ): void => {
      const slots = [] as ChartSlot[];
      data.series.forEach((series: ChartSeries, s: number) => {
        slots.push({
          index: s,
          series: series,
          targetCol: s + 1,
          type: 'extendedSeries',
        });
      });

      new ModifyChart(chart, workbook, data, slots).modifyExtended();

      // XmlHelper.dump(chart);
      // XmlHelper.dump(workbook.table)
    };

  static setAxisRange =
    (range: ChartAxisRange) =>
    (chart: XmlDocument): void => {
      const axis = chart.getElementsByTagName('c:valAx')[range.axisIndex || 0];
      if (!axis) return;

      ModifyChartHelper.setAxisAttribute(axis, 'c:majorUnit', range.majorUnit);
      ModifyChartHelper.setAxisAttribute(axis, 'c:minorUnit', range.minorUnit);
      ModifyChartHelper.setAxisAttribute(
        axis,
        'c:numFmt',
        range.formatCode,
        'formatCode',
      );
      ModifyChartHelper.setAxisAttribute(
        axis,
        'c:numFmt',
        range.sourceLinked,
        'sourceLinked',
      );

      const scaling = axis.getElementsByTagName('c:scaling')[0];

      ModifyChartHelper.setAxisAttribute(scaling, 'c:min', range.min);
      ModifyChartHelper.setAxisAttribute(scaling, 'c:max', range.max);
    };

  static setAxisAttribute = (
    element: XmlElement,
    tag: string,
    value: string | number | boolean,
    attribute?: string,
  ) => {
    if (value === undefined || !element) return;
    const target = element.getElementsByTagName(tag);
    if (target.length > 0) {
      attribute = attribute || 'val';
      if (typeof value === 'boolean') {
        ModifyXmlHelper.booleanAttribute(attribute, value)(target[0]);
      } else {
        ModifyXmlHelper.attribute(attribute, value)(target[0]);
      }
    }
  };
  static setLabelHidden = () =>
  (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
  ): void => {
      this.setLabelArea(
        {
          w:0.0, h:0.0, x:0.0, y:0.0
        })(element, chart, workbook);
  }
  static setLabelArea = (legendArea: ChartPlotArea) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
    ): void => {
      const modifyXmlHelper = new ModifyXmlHelper(chart);
      modifyXmlHelper.modify({
        'c:legend': {
          children: {
            'c:manualLayout': {
              children: {
                'c:w': {
                  modify: [
                    ModifyXmlHelper.attribute('val', legendArea.w),
                  ],
                },
                'c:h': {
                  modify: [
                    ModifyXmlHelper.attribute('val', legendArea.h),
                  ],
                },
                'c:x': {
                  modify: [
                    ModifyXmlHelper.attribute('val', legendArea.x),
                  ],
                },
                'c:y': {
                  modify: [
                    ModifyXmlHelper.attribute('val', legendArea.y),
                  ],
                },
              },
            },
          },
        },
      });
      XmlHelper.dump(
        chart
          .getElementsByTagName('c:legendPos')[0]
      );
    };
  /**
   * Set plot area size.
   */
  static setPlotArea =
    (plotArea: ChartPlotArea) =>
    (
      element: XmlDocument | XmlElement,
      chart?: XmlDocument,
      workbook?: Workbook,
    ): void => {
      // Each chart has a separate chart xml file. It is required
      // to alter everything that's "inside" the chart, e.g. data, legend,
      // axis... and: plot area

      // ModifyXmlHelper class provides a lot of functions to access
      // and edit xml elements.
      const modifyXmlHelper = new ModifyXmlHelper(chart);

      // We need to locate the required xml elements and target them
      // with ModifyXmlHelper's help.
      // We can therefore log the entire chart.xml to console:
      // XmlHelper.dump(chart);

      modifyXmlHelper.modify({
        'c:plotArea': {
          children: {
            'c:manualLayout': {
              children: {
                'c:w': {
                  // Finally, we attach ModifyCallbacks to all
                  // matching elements
                  modify: [
                    ModifyXmlHelper.attribute('val', plotArea.w),
                    // ...
                  ],
                },
                'c:h': {
                  modify: [
                    ModifyXmlHelper.attribute('val', plotArea.h),
                  ],
                },
                'c:x': {
                  modify: [
                    ModifyXmlHelper.attribute('val', plotArea.x),
                  ],
                },
                'c:y': {
                  modify: [
                    ModifyXmlHelper.attribute('val', plotArea.y),
                  ],
                },
              },
            },
          },
        },
      });

      // We can dump the target node and see if our modification
      // took effect.
      XmlHelper.dump(
        chart
          .getElementsByTagName('c:plotArea')[0]
          .getElementsByTagName('c:manualLayout')[0],
      );

      // You can also take a look at element xml, which is a child node
      // of current slide. It holds general shape properties, but no
      // data or so.
      // XmlHelper.dump(element);

      // Rough ones might also want to look inside the linked workbook.
      // It is located inside an extra xlsx file. We don't need this
      // for now.
      // XmlHelper.dump(workbook.table)
      // XmlHelper.dump(workbook.sheet)
    };
}
