import Automizer from './automizer';

export default Automizer;

import {
  setAttribute,
  dump,
} from './helper/modify-helper';

import {
  setSolidFill,
  setText,
  setPosition,
} from './helper/modify-shape-helper';

import {
  setTableData
} from './helper/modify-table-helper';

import {
  setChartData,
  setChartVerticalLines,
  setChartBubbles,
} from './helper/modify-chart-helper';

export const modify = {
  setSolidFill,
  setText,
  setPosition,
  setAttribute,
  setChartData,
  setChartVerticalLines,
  setChartBubbles,
  setTableData,
  dump,
};
