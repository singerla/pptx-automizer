export type TableRow = {
  label?: string;
  values: (string | number)[];
};

export type TableData = {
  header?: TableRow | TableRow[];
  body?: TableRow[];
  footer?: TableRow | TableRow[];
};

export type ModifyTableParams = {
  adjustWidth: boolean;
  adjustHeight: boolean;
};
