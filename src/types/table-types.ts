import { Color, TextStyle } from './modify-types';

export type TableRow = {
  label?: string;
  values: (string | number)[];
  styles?: (null | TableRowStyle)[];
};

export type TableRowStyle = TextStyle & {
  background?: Color;
};

export type TableData = {
  header?: TableRow | TableRow[];
  body?: TableRow[];
  footer?: TableRow | TableRow[];
};

export type ModifyTableParams = {
  adjustWidth: boolean;
  adjustHeight: boolean;
  setHeight?: number;
  setWidth?: number;
};
