import Automizer from './automizer';

import ModifyHelper, { CmToDxa, DxaToCm } from './helper/modify-helper';
import ModifyShapeHelper from './helper/modify-shape-helper';
import ModifyTableHelper from './helper/modify-table-helper';
import ModifyChartHelper from './helper/modify-chart-helper';
import ModifyHyperlinkHelper from './helper/modify-hyperlink-helper';
import { TemplateInfo, XmlDocument, XmlElement } from './types/xml-types';
import {
  AutomizerFile,
  AutomizerParams,
  AutomizerSummary,
  ShapeModificationCallback,
  StatusTracker,
} from './types/types';
import {
  ModifyTableParams,
  TableData,
  TableRow,
  TableRowStyle,
} from './types/table-types';
import {
  ChartBubble,
  ChartCategory,
  ChartData,
  ChartSeries,
  ChartValueStyle,
} from './types/chart-types';
import {
  Border,
  Color,
  ImageStyle,
  ReplaceText,
  TextStyle,
} from './types/modify-types';
import { ShapeCoordinates, ShapeValueStyle } from './types/shape-types';
import { XmlHelper } from './helper/xml-helper';
import ModifyTextHelper from './helper/modify-text-helper';
import ModifyColorHelper from './helper/modify-color-helper';

import { ITemplate } from './interfaces/itemplate';
import { IMaster } from './interfaces/imaster';
import { ISlide } from './interfaces/islide';
import { IPptxGenJSSlide } from './interfaces/ipptxgenjs-slide';
import ModifyImageHelper from './helper/modify-image-helper';
import { LabelPosition } from './enums/chart-type';

const dump = ModifyHelper.dump;
const dumpChart = ModifyHelper.dumpChart;
const setAttribute = ModifyHelper.setAttribute;

const setSolidFill = ModifyShapeHelper.setSolidFill;
const setText = ModifyShapeHelper.setText;
const setBulletList = ModifyShapeHelper.setBulletList;
const replaceText = ModifyShapeHelper.replaceText;
const setPosition = ModifyShapeHelper.setPosition;
const updatePosition = ModifyShapeHelper.updatePosition;
const rotateShape = ModifyShapeHelper.rotate;

const setTableData = ModifyTableHelper.setTableData;
const adjustHeight = ModifyTableHelper.adjustHeight;
const adjustWidth = ModifyTableHelper.adjustWidth;
const setTable = ModifyTableHelper.setTable;
const updateColumnWidth = ModifyTableHelper.updateColumnWidth;
const updateRowHeight = ModifyTableHelper.updateRowHeight;

const setRelationTarget = ModifyImageHelper.setRelationTarget;
const setDuotoneFill = ModifyImageHelper.setDuotoneFill;

const setChartData = ModifyChartHelper.setChartData;
const setExtendedChartData = ModifyChartHelper.setExtendedChartData;
const setChartVerticalLines = ModifyChartHelper.setChartVerticalLines;
const setChartScatter = ModifyChartHelper.setChartScatter;
const setChartBubbles = ModifyChartHelper.setChartBubbles;
const setChartCombo = ModifyChartHelper.setChartCombo;
const setAxisRange = ModifyChartHelper.setAxisRange;
const setPlotArea = ModifyChartHelper.setPlotArea;
const setLegendPosition = ModifyChartHelper.setLegendPosition;
const removeChartLegend = ModifyChartHelper.removeChartLegend;
const minimizeChartLegend = ModifyChartHelper.minimizeChartLegend;
const setWaterFallColumnTotalToLast =
  ModifyChartHelper.setWaterFallColumnTotalToLast;
const setChartTitle = ModifyChartHelper.setChartTitle;
const setDataLabelAttributes = ModifyChartHelper.setDataLabelAttributes;
const readWorkbookData = ModifyChartHelper.readWorkbookData;
const readChartInfo = ModifyChartHelper.readChartInfo;

const setHyperlinkTarget = ModifyHyperlinkHelper.setHyperlinkTarget;
const addHyperlink = ModifyHyperlinkHelper.addHyperlink;
const removeHyperlink = ModifyHyperlinkHelper.removeHyperlink;

export type {
  ChartBubble,
  ChartCategory,
  ChartSeries,
  TableData,
  TableRow,
  TextStyle,
  ImageStyle,
  ModifyTableParams,
  TableRowStyle,
  ChartData,
  ChartValueStyle,
  Color,
  Border,
  ShapeValueStyle,
  ShapeCoordinates,
  ReplaceText,
};

export {
  Automizer,
  XmlHelper,
  ModifyHelper,
  ModifyShapeHelper,
  ModifyTableHelper,
  ModifyChartHelper,
  ModifyTextHelper,
  ModifyColorHelper,
  ModifyImageHelper,
  LabelPosition,
};

export const modify = {
  dump,
  dumpChart,
  setAttribute,
  setSolidFill,
  setText,
  setBulletList,
  replaceText,
  setPosition,
  updatePosition,
  rotateShape,
  setTableData,
  adjustHeight,
  adjustWidth,
  updateColumnWidth,
  updateRowHeight,
  setTable,
  setRelationTarget,
  setDuotoneFill,
  setChartData,
  setAxisRange,
  setExtendedChartData,
  setChartVerticalLines,
  setChartScatter,
  setChartCombo,
  setChartBubbles,
  setPlotArea,
  setLegendPosition,
  removeChartLegend,
  minimizeChartLegend,
  setWaterFallColumnTotalToLast,
  setChartTitle,
  setDataLabelAttributes,
  setHyperlinkTarget,
  addHyperlink,
  removeHyperlink,
};

export const read = {
  readWorkbookData,
  readChartInfo,
};

export { CmToDxa, DxaToCm };

export type {
  AutomizerFile,
  AutomizerParams,
  AutomizerSummary,
  TemplateInfo,
  StatusTracker,
  ShapeModificationCallback,
  XmlElement,
  XmlDocument,
};

export { ITemplate, IMaster, ISlide, IPptxGenJSSlide };
export default Automizer;
