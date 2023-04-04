import Automizer from './automizer';

import ModifyHelper from './helper/modify-helper';
import ModifyShapeHelper from './helper/modify-shape-helper';
import ModifyTableHelper from './helper/modify-table-helper';
import ModifyChartHelper from './helper/modify-chart-helper';
import { TemplateInfo } from './types/xml-types';
import {
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
import { Color, ReplaceText, TextStyle } from './types/modify-types';
import { ShapeCoordinates, ShapeValueStyle } from './types/shape-types';
import { XmlHelper } from './helper/xml-helper';
import ModifyTextHelper from './helper/modify-text-helper';
import ModifyColorHelper from './helper/modify-color-helper';

const dump = ModifyHelper.dump;
const dumpChart = ModifyHelper.dumpChart;
const setAttribute = ModifyHelper.setAttribute;

const setSolidFill = ModifyShapeHelper.setSolidFill;
const setText = ModifyShapeHelper.setText;
const replaceText = ModifyShapeHelper.replaceText;
const setPosition = ModifyShapeHelper.setPosition;
const updatePosition = ModifyShapeHelper.updatePosition;

const setTableData = ModifyTableHelper.setTableData;
const adjustHeight = ModifyTableHelper.adjustHeight;
const adjustWidth = ModifyTableHelper.adjustWidth;
const setTable = ModifyTableHelper.setTable;

const setChartData = ModifyChartHelper.setChartData;
const setExtendedChartData = ModifyChartHelper.setExtendedChartData;
const setChartVerticalLines = ModifyChartHelper.setChartVerticalLines;
const setChartScatter = ModifyChartHelper.setChartScatter;
const setChartBubbles = ModifyChartHelper.setChartBubbles;
const setChartCombo = ModifyChartHelper.setChartCombo;

export type {
  ChartBubble,
  ChartCategory,
  ChartSeries,
  TableData,
  TableRow,
  TextStyle,
  ModifyTableParams,
  TableRowStyle,
  ChartData,
  ChartValueStyle,
  Color,
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
};

export const modify = {
  dump,
  dumpChart,
  setAttribute,
  setSolidFill,
  setText,
  replaceText,
  setPosition,
  updatePosition,
  setTableData,
  adjustHeight,
  adjustWidth,
  setTable,
  setChartData,
  setExtendedChartData,
  setChartVerticalLines,
  setChartScatter,
  setChartCombo,
  setChartBubbles,
};
export type {
  AutomizerSummary,
  TemplateInfo,
  StatusTracker,
  ShapeModificationCallback,
};
export default Automizer;
