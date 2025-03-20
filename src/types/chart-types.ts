import { Color, ModificationTags } from './modify-types';
import { ShapeCoordinates } from './shape-types';
import { LabelPosition } from '../enums/chart-type';

export type ChartPointValue = null | number;
export type ChartValueStyle = {
  color?: Color;
  background?: Color;
  marker?: {
    color?: Color;
  };
  border?: {
    color?: Color;
    weight?: number;
  };
  label?: {
    color?: Color;
    isBold?: boolean;
    size?: number;
  } & ChartDataLabelAttributes;
  gradient?: {
    color: Color;
    index: number;
  }[];
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
  styles?: (ChartValueStyle | null)[];
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
  (point: number | ChartPoint | ChartBubble, category?: ChartCategory): number;
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
export type ChartAxisRange = {
  axisIndex?: number;
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  formatCode?: string;
  sourceLinked?: boolean;
};

export type ChartSeriesDataLabelAttributes = {
  applyToSeries?: number;
} & ChartDataLabelAttributes;

export type ChartDataLabelAttributes = {
  dLblPos?: LabelPosition;
  showLegendKey?: boolean;
  showVal?: boolean;
  showCatName?: boolean;
  showSerName?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  showLeaderLines?: boolean;
  solidFill?: Color;
};
// Elements inside a chart (e.g. a legend) require shares as coordinates.
// E.g. "w: 0.5" means "half of chart width"
export type ChartElementCoordinateShares = ShapeCoordinates;
