import { Color, Border, TextStyle } from './modify-types';
import { XmlElement } from './xml-types';

export type TableRow = {
  label?: string;
  values: (string | number)[];
  styles?: (null | TableRowStyle)[];
};

export type TableRowStyle = TextStyle & {
  background?: Color;
  border?: Border[];
};

export type TableData = {
  header?: TableRow | TableRow[];
  body?: TableRow[];
  footer?: TableRow | TableRow[];
};

export type ModifyTableParams = {
  adjustWidth?: boolean;
  adjustHeight?: boolean;
  setHeight?: number;
  setWidth?: number;
  expand?: ModifyTableExpand[];
};

export type ModifyTableExpand = {
  tag: string;
  mode: 'row' | 'column';
  count: number;
};

export type TableInfo = {
  row: number;
  column: number;
  rowXml: XmlElement;
  columnXml: XmlElement;
  text: string[];
  textContent: string;
  gridSpan: number;
  hMerge: number;
};
