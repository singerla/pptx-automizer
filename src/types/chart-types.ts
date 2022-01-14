import {Color, ModificationTags} from './modify-types';

export type ChartPointValue = null|number
export type ChartValueStyle = {
  color?: Color;
  marker?: {
    color?: Color;
  }
};
export type ChartPoint = {
  x: ChartPointValue;
  y: ChartPointValue;
};
export type ChartBubble = {
  x: ChartPointValue;
  y: ChartPointValue;
  size: number;
};
export type ChartSeries = {
  label: string;
  style?: ChartValueStyle;
};
export type ChartCategory = {
  label: string;
  y?: ChartPointValue;
  values: (ChartPointValue | ChartPoint | ChartBubble)[];
  styles?: ChartValueStyle[]
};
export type ChartColumn = {
  series?: number;
  label: string;
  worksheet: (
    point: ChartPointValue | ChartPoint | ChartBubble,
    r: number,
    category: ChartCategory,
  ) => void;
  chart?: (
    point: ChartPointValue | ChartPoint | ChartBubble,
    r: number,
    category: ChartCategory,
  ) => ModificationTags;
  isStrRef?: boolean;
  modTags?: ModificationTags;
};
export type ChartData = {
  series: ChartSeries[];
  categories: ChartCategory[];
};
export type ChartDataMapper = {
  (
    point: number | ChartPoint | ChartBubble,
    category?: ChartCategory,
  ): number;
};
export type ChartSlot = {
  label?: string;
  mapData?: ChartDataMapper;
  series?: ChartSeries;
  index?: number;
  targetCol: number;
  targetYCol?: number;
  type?: string;
  tag?: string;
  isStrRef?: boolean;
};
