import { ModificationTags } from './modify-types';

export type ChartValue = {
  value: number;
};
export type ChartPoint = {
  x: number;
  y: number;
};
export type ChartBubble = {
  x: number;
  y: number;
  size: number;
};
export type ChartSeries = {
  label: string;
};
export type ChartCategory = {
  label: string;
  y?: number;
  values: number[] | ChartValue[] | ChartPoint[] | ChartBubble[];
};
export type ChartColumn = {
  series?: number;
  label: string;
  worksheet: (
    point: number | ChartPoint | ChartBubble | ChartValue,
    r: number,
    category: ChartCategory,
  ) => void;
  chart?: (
    point: number | ChartPoint | ChartBubble | ChartValue,
    r: number,
    category: ChartCategory,
  ) => ModificationTags;
  isStrRef?: boolean;
};
export type ChartData = {
  series: ChartSeries[];
  categories: ChartCategory[];
};
export type ChartDataMapper = {
  (
    point: number | ChartPoint | ChartBubble | ChartValue,
    category?: ChartCategory,
  ): number;
};
export type ChartSlot = {
  label?: string;
  mapData?: ChartDataMapper;
  series?: ChartSeries;
  index?: number;
  targetCol: number;
  type?: string;
  tag?: string;
  isStrRef?: boolean;
};
