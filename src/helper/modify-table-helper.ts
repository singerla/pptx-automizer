import { XmlHelper } from './xml-helper';
import { TableData } from '../types/table-types';

export const setTableData = (data: TableData) => (
  element: XMLDocument | Document | Element,
): void => {
  XmlHelper.dump(element);
};
