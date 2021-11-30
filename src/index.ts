import Automizer from './automizer';

export default Automizer;

import ModifyHelper from './helper/modify-helper';
const dump = ModifyHelper.dump;
const setAttribute = ModifyHelper.setAttribute;

import ModifyShapeHelper from './helper/modify-shape-helper';
const setSolidFill = ModifyShapeHelper.setSolidFill;
const setText = ModifyShapeHelper.setText;
const replaceText = ModifyShapeHelper.replaceText;
const setPosition = ModifyShapeHelper.setPosition;

import ModifyTableHelper from './helper/modify-table-helper';
const setTableData = ModifyTableHelper.setTableData;
const adjustHeight = ModifyTableHelper.adjustHeight;
const adjustWidth = ModifyTableHelper.adjustWidth;
const setTable = ModifyTableHelper.setTable;

import ModifyChartHelper from './helper/modify-chart-helper';
const setChartData = ModifyChartHelper.setChartData;
const setChartVerticalLines = ModifyChartHelper.setChartVerticalLines;
const setChartBubbles = ModifyChartHelper.setChartBubbles;

import { AutomizerSummary } from './types/types';
export type { AutomizerSummary };

import { ModifyTableParams } from './types/table-types';
import { ChartData } from './types/chart-types';

export type { ModifyTableParams, ChartData };

export {
  Automizer,
  ModifyHelper,
  ModifyShapeHelper,
  ModifyTableHelper,
  ModifyChartHelper,
};

export const modify = {
  dump,
  setAttribute,
  setSolidFill,
  setText,
  replaceText,
  setPosition,
  setTableData,
  adjustHeight,
  adjustWidth,
  setTable,
  setChartData,
  setChartVerticalLines,
  setChartBubbles,
};
