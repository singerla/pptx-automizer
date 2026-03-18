import { ModifyChart } from '../modify/modify-chart';
import { ChartModificationCallback, Workbook } from '../types/types';
import {
  ChartAxisRange,
  ChartBubble,
  ChartCategory,
  ChartData,
  ChartElementCoordinateShares,
  ChartPoint, ChartPointValue,
  ChartSeries,
  ChartSeriesDataLabelAttributes,
  ChartSlot,
  ChartValueStyle,
} from '../types/chart-types';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import ModifyColorHelper from './modify-color-helper';

export default class ModifyChartHelper {
  /**
   * Set chart data to modify default chart types.
   * See `__tests__/modify-existing-chart.test.js`
   */
  static setChartData =
    (data: ChartData): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
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
    (data: ChartData): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
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

      // ModifyChartHelper.setAxisRange({
      //   axisIndex: 0,
      //   min: 0,
      //   max: data.categories.length,
      // })(element, chart);
    };

  /**
   * Set chart data to modify scatter charts.
   * See `__tests__/modify-chart-scatter.test.js`
   */
  static setChartScatter =
    (data: ChartData): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
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
    (data: ChartData): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
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

      ModifyChartHelper.setAxisRange({
        axisIndex: 1,
        min: 0,
        max: data.categories.length,
      })(element, chart);
    };

  /**
   * Set chart data to modify bubble charts.
   * See `__tests__/modify-chart-bubbles.test.js`
   */
  static setChartBubbles =
    (data: ChartData): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
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
    (data: ChartData): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
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

  /**
   * Read chart workbook data
   * See `__tests__/read-chart-data.test.js`
   */
  static readWorkbookData =
    (data: any): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
      const getSharedString = (index: number): string => {
        return workbook.sharedStrings.getElementsByTagName('si').item(index)
          ?.textContent;
      };

      const parseCell = (cell: XmlElement): string | number => {
        const type = cell.getAttribute('t');
        const cellValue = cell.getElementsByTagName('v').item(0).textContent;
        if (type === 's') {
          return getSharedString(Number(cellValue));
        } else {
          return Number(cellValue);
        }
      };

      const rows = workbook.sheet.getElementsByTagName('row');
      for (let r = 0; r < rows.length; r++) {
        const row = rows.item(r);
        const columns = row.getElementsByTagName('c');
        const rowData = [];
        for (let c = 0; c < columns.length; c++) {
          rowData.push(parseCell(columns.item(c)));
        }
        data.push(rowData);
      }
    };

  /**
   * Read chart info
   * See `__tests__/read-chart-data.test.js`
   */
  static readChartInfo =
    (info: any): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument, workbook?: Workbook): void => {
      const series = chart.getElementsByTagName('c:ser');
      XmlHelper.modifyCollection(series, (tmpSeries: XmlElement, s: number) => {
        const solidFill = tmpSeries.getElementsByTagName('a:solidFill').item(0);
        if (!solidFill) {
          return;
        }

        const schemeClr = solidFill.getElementsByTagName('a:schemeClr').item(0);
        const srgbClr = solidFill.getElementsByTagName('a:srgbClr').item(0);
        const colorElement = schemeClr ? schemeClr : srgbClr;
        info.series.push({
          seriesId: s,
          colorType: colorElement.tagName,
          colorValue: colorElement.getAttribute('val'),
        });
      });

      const chartTagName = series.item(0).parentNode.nodeName;
      info.chartType = chartTagName?.split(':')[1];
    };

  /**
   * Set range and format for chart axis.
   * Please notice: It will only work if the value to update is not set to
   * "Auto" in powerpoint. Only manually scaled min/max can be altered by this.
   * See `__tests__/modify-chart-axis.test.js`
   */
  static setAxisRange =
    (range: ChartAxisRange): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument): void => {
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

  /**
   * Set legend coordinates to zero. Could be advantageous for pptx users to
   * be able to maximize a legend easily. Legend will still be selectible for
   * a user.
   */
  static minimizeChartLegend =
    (): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument, workbook?: Workbook): void => {
      this.setLegendPosition({
        w: 0.0,
        h: 0.0,
        x: 0.0,
        y: 0.0,
      })(element, chart, workbook);
    };

  /**
   * Completely remove a chart legend. Please notice: This will trigger
   * PowerPoint to automatically maximize chart space.
   */
  static removeChartLegend =
    (): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument): void => {
      if (chart.getElementsByTagName('c:legend')) {
        XmlHelper.remove(chart.getElementsByTagName('c:legend')[0]);
      }
    };

  /**
   * Update the coordinates of a chart legend.
   * legendArea coordinates are shares of chart coordinates, e.g.
   * "w: 0.5" means "half of chart width"
   * @param legendArea
   */
  static setLegendPosition =
    (legendArea: ChartElementCoordinateShares): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument): void => {
      const modifyXmlHelper = new ModifyXmlHelper(chart);
      modifyXmlHelper.modify({
        'c:legend': {
          children: {
            'c:manualLayout': {
              children: {
                'c:w': {
                  modify: [ModifyXmlHelper.attribute('val', legendArea.w)],
                },
                'c:h': {
                  modify: [ModifyXmlHelper.attribute('val', legendArea.h)],
                },
                'c:x': {
                  modify: [ModifyXmlHelper.attribute('val', legendArea.x)],
                },
                'c:y': {
                  modify: [ModifyXmlHelper.attribute('val', legendArea.y)],
                },
              },
            },
          },
        },
      });
      // XmlHelper.dump(chart.getElementsByTagName('c:legendPos')[0]);
    };

  /**
   * Set the plot area coordinates of a chart.
   *
   * This modifier requires a 'c:manualLayout' element. It will only appear if
   * plot area coordinates are edited manually in ppt before. Recently fresh
   * created charts will not have a manualLayout by default.
   *
   * This is especially useful if you have problems with edgy elements on a
   * chart area that do not fit into the given space, e.g. when having a lot
   * of data labels. You can increase the chart and decrease the plot area
   * to create a margin.
   *
   * plotArea coordinates are shares of chart coordinates, e.g.
   * "w: 0.5" means "half of chart width"
   *
   * @param plotArea
   */
  static setPlotArea =
    (plotArea: ChartElementCoordinateShares): ChartModificationCallback =>
    (element: XmlElement, chart?: XmlDocument): void => {
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

      // There needs to be a 'c:manualLayout' element. This will only appear if
      // a plot area was edited manually in ppt before. Recently fresh created
      // charts will not have a manualLayout by default.
      if (
        !chart
          .getElementsByTagName('c:plotArea')[0]
          .getElementsByTagName('c:manualLayout')[0]
      ) {
        console.error("Can't update plot area. No c:manualLayout found.");
        return;
      }

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
                  modify: [ModifyXmlHelper.attribute('val', plotArea.h)],
                },
                'c:x': {
                  modify: [ModifyXmlHelper.attribute('val', plotArea.x)],
                },
                'c:y': {
                  modify: [ModifyXmlHelper.attribute('val', plotArea.y)],
                },
              },
            },
          },
        },
      });

      // We can dump the target node and see if our modification
      // took effect.
      // XmlHelper.dump(
      //   chart
      //     .getElementsByTagName('c:plotArea')[0]
      //     .getElementsByTagName('c:manualLayout')[0],
      // );

      // You can also take a look at element xml, which is a child node
      // of current slide. It holds general shape properties, but no
      // data or so.
      // XmlHelper.dump(chart);

      // Rough ones might also want to look inside the linked workbook.
      // It is located inside an extra xlsx file. We don't need this
      // for now.
      // XmlHelper.dump(workbook.table)
      // XmlHelper.dump(workbook.sheet)
    };

  /**
   * Set a waterfall Total column to last
   * you may also optionally specify a different index.
   @param TotalColumnIDX
   *
   */
  static setWaterFallColumnTotalToLast =
    (TotalColumnIDX?: number): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument): void => {
      const plotArea = chart.getElementsByTagName('cx:plotArea')[0];
      const subTotals = plotArea
        ?.getElementsByTagName('cx:layoutPr')[0]
        ?.getElementsByTagName('cx:subtotals')[0];

      if (subTotals) {
        if (!TotalColumnIDX) {
          const GetTotalPoints = chart
            .getElementsByTagName('cx:chartData')[0]
            ?.getElementsByTagName('cx:data')[0]
            ?.getElementsByTagName('cx:strDim')[0]
            ?.getElementsByTagName('cx:lvl')[0]
            ?.getAttribute('ptCount');
          if (GetTotalPoints) {
            TotalColumnIDX = Number(GetTotalPoints) - 1;
          }
        }
        if (TotalColumnIDX !== undefined) {
          const stIndexes = Array.from(
            subTotals.getElementsByTagName('cx:idx'),
          );
          stIndexes.forEach((sTValue, index) => {
            ModifyXmlHelper.attribute(
              'val',
              TotalColumnIDX.toString(),
            )(sTValue);
            if (index > 0) {
              subTotals.removeChild(sTValue);
            }
          });
        }
      }
    };

  /**
   * Set the title of a chart. This requires an already existing, manually edited chart title.
   @param newTitle
   *
   */
  static setChartTitle =
    (newTitle: string): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument): void => {
      const chartTitle = chart.getElementsByTagName('c:title').item(0);
      const chartTitleText = chartTitle?.getElementsByTagName('a:t').item(0);
      if (chartTitleText) {
        chartTitleText.textContent = XmlHelper.sanitizeText(newTitle);
      }
    };

  static setPointLabelSuffix = (
    element: XmlElement,
    idx: number,
    labelStyle: ChartValueStyle['label'],
  ): void => {
    const doc = element.ownerDocument;

    // Convert c:txPr to c:tx > c:rich so PowerPoint renders text runs.
    const txPr = element.getElementsByTagName('c:txPr')[0];
    if (txPr) {
      const tx = doc.createElement('c:tx');
      const rich = doc.createElement('c:rich');

      while (txPr.firstChild) {
        rich.appendChild(txPr.firstChild);
      }

      tx.appendChild(rich);
      txPr.parentNode.insertBefore(tx, txPr);
      txPr.parentNode.removeChild(txPr);
    }

    const paragraphs = element.getElementsByTagName('a:p');
    if (!paragraphs.length) return;
    const p = paragraphs[0];

    // Remove any existing a:r and a:fld elements that may have been
    // carried over from cloning a previously modified c:dLbl.
    const existingRuns = Array.from(p.getElementsByTagName('a:r'));
    existingRuns.forEach((run) => run.parentNode.removeChild(run));
    const existingFlds = Array.from(p.getElementsByTagName('a:fld'));
    existingFlds.forEach((fld) => fld.parentNode.removeChild(fld));

    // Add a:fld element to display the chart value
    const endParaRPr = p.getElementsByTagName('a:endParaRPr')[0];
    const defRPr = p.getElementsByTagName('a:defRPr')[0];

    const fld = doc.createElement('a:fld');
    fld.setAttribute('type', 'VALUE');
    fld.setAttribute(
      'id',
      '{AABBCCDD-1234-5678-9012-' + String(idx).padStart(12, '0') + '}',
    );

    const fldRPr = doc.createElement('a:rPr');
    fldRPr.setAttribute('lang', 'en-US');
    if (defRPr) {
      Array.from(defRPr.childNodes).forEach((child) => {
        fldRPr.appendChild(child.cloneNode(true));
      });
      Array.from(defRPr.attributes).forEach((attr) => {
        if (attr.name !== 'sz' && attr.name !== 'b') {
          fldRPr.setAttribute(attr.name, attr.value);
        }
      });
    }
    fld.appendChild(fldRPr);

    const fldT = doc.createElement('a:t');
    fldT.textContent = '[VALUE]';
    fld.appendChild(fldT);

    if (endParaRPr) {
      p.insertBefore(fld, endParaRPr);
    } else {
      p.appendChild(fld);
    }

    // Add the suffix a:r
    const r = doc.createElement('a:r');
    const rPr = doc.createElement('a:rPr');
    rPr.setAttribute('lang', 'en-US');
    rPr.setAttribute('dirty', '0');

    if (labelStyle.suffix.color) {
      const color = { ...labelStyle.suffix.color };
      if (color.value.indexOf('#') === 0) {
        color.value = color.value.replace('#', '');
      }
      const solidFill = doc.createElement('a:solidFill');
      const colorEl = doc.createElement('a:' + color.type);
      colorEl.setAttribute('val', color.value);
      solidFill.appendChild(colorEl);
      rPr.appendChild(solidFill);
    }

    r.appendChild(rPr);
    const t = doc.createElement('a:t');
    t.textContent = labelStyle.suffix.text;
    r.appendChild(t);

    const endParaRPr2 = p.getElementsByTagName('a:endParaRPr')[0];
    if (endParaRPr2) {
      p.insertBefore(r, endParaRPr2);
    } else {
      p.appendChild(r);
    }

    // Prevent line breaks in the data label by setting wrap="none"
    const bodyPrs = element.getElementsByTagName('a:bodyPr');
    if (bodyPrs.length) {
      bodyPrs[0].setAttribute('wrap', 'none');
    }

    // Preserve the number format of the datalabel if it exists in the series or the current label.
    // If it's missing in the current c:dLbl, look into the parent c:dLbls.
    if (!element.getElementsByTagName('c:numFmt').length) {
      const dLbls = element.parentNode;
      if (dLbls && dLbls.nodeName === 'c:dLbls') {
        const numFmt = (dLbls as XmlElement).getElementsByTagName('c:numFmt')[0];
        if (numFmt) {
          element.insertBefore(numFmt.cloneNode(true), element.firstChild);
        }
      }
    }
  };

  /**
   * Specify a format for DataLabels
   @param dataLabel
   *
   */
  static setDataLabelAttributes =
    (dataLabel: ChartSeriesDataLabelAttributes): ChartModificationCallback =>
    (element: XmlElement, chart: XmlDocument): void => {
      const modifyXmlHelper = new ModifyXmlHelper(chart);
      const applyToSeries =
        typeof dataLabel.applyToSeries === 'number'
          ? {
              index: dataLabel.applyToSeries,
            }
          : {
              all: true,
            };

      modifyXmlHelper.modify({
        'c:ser': {
          ...applyToSeries,
          children: {
            'c:dLbls': {
              children:
                ModifyChartHelper.setDataPointLabelAttributes(dataLabel),
            },
          },
        },
      });
    };

  static setDataPointLabelAttributes = (
    dataLabel: ChartSeriesDataLabelAttributes,
  ) => {
    return {
      'c:spPr': {
        modify: [ModifyColorHelper.solidFill(dataLabel.solidFill)],
      },
      'c:numFmt': {
        modify: [
          ModifyXmlHelper.attribute('formatCode', dataLabel.formatCode),
          ModifyXmlHelper.booleanAttribute(
            'sourceLinked',
            dataLabel.sourceLinked,
          ),
        ],
      },
      'c:dLblPos': {
        modify: [ModifyXmlHelper.attribute('val', dataLabel.dLblPos)],
      },
      'c:showLegendKey': {
        modify: [
          ModifyXmlHelper.booleanAttribute('val', dataLabel.showLegendKey),
        ],
      },
      'c:showVal': {
        modify: [ModifyXmlHelper.booleanAttribute('val', dataLabel.showVal)],
      },
      'c:showCatName': {
        modify: [
          ModifyXmlHelper.booleanAttribute('val', dataLabel.showCatName),
        ],
      },
      'c:showSerName': {
        modify: [
          ModifyXmlHelper.booleanAttribute('val', dataLabel.showSerName),
        ],
      },
      'c:showPercent': {
        modify: [
          ModifyXmlHelper.booleanAttribute('val', dataLabel.showPercent),
        ],
      },
      'c:showBubbleSize': {
        modify: [
          ModifyXmlHelper.booleanAttribute('val', dataLabel.showBubbleSize),
        ],
      },
      'c:showLeaderLines': {
        modify: [
          ModifyXmlHelper.booleanAttribute('val', dataLabel.showLeaderLines),
        ],
      },
    };
  };

  // Prevent a "null" from being written to the xml file.
  static parseCellValue(value: string | ChartPointValue): string {
    if (value === null) return '';

    return String(value);
  }
}
