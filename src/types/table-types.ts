export type TableCell = {
  value: number | string;
};

export type TableRow = {
  label: string;
  values: TableCell['value'][] | TableCell[];
};

export type TableData = {
  header?: TableRow | TableRow[];
  body?: TableRow[];
  footer?: TableRow | TableRow[];
  [index: number]: TableRow['values'];
};
