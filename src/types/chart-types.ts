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
    ctx: any,
    point: number | ChartPoint | ChartBubble | ChartValue,
    r: number,
    category: ChartCategory,
  ) => any;
  chart?: (
    ctx: any,
    point: number | ChartPoint | ChartBubble | ChartValue,
    r: number,
    category: ChartCategory,
  ) => any;
  isStrRef?: boolean;
};
export type ChartData = {
  series: ChartSeries[];
  categories: ChartCategory[];
};
export type ModificationPatternModifier = {
  (element: Element);
};
export type ModificationPattern = {
  index?: number;
  children?: ModificationPatternChildren;
  modify?: ModificationPatternModifier | ModificationPatternModifier[];
};
export type ModificationPatternChildren = {
  [tag: string]: ModificationPattern;
};
