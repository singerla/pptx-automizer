import Automizer from './automizer';

export default Automizer;

import ModifyHelper from './helper/modify-helper';
const dump = ModifyHelper.dump;
const setAttribute = ModifyHelper.setAttribute;

import ModifyShapeHelper from './helper/modify-shape-helper';
const setSolidFill = ModifyShapeHelper.setSolidFill;
const setText = ModifyShapeHelper.setText;
const setPosition = ModifyShapeHelper.setPosition;

import ModifyTableHelper from './helper/modify-table-helper';
const setTableData = ModifyTableHelper.setTableData;

import ModifyChartHelper from './helper/modify-chart-helper';
const setChartData = ModifyChartHelper.setChartData;
const setChartVerticalLines = ModifyChartHelper.setChartVerticalLines;
const setChartBubbles = ModifyChartHelper.setChartBubbles;

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
  setPosition,
  setTableData,
  setChartData,
  setChartVerticalLines,
  setChartBubbles,
};
