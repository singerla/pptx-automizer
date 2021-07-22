import { XmlHelper } from './xml-helper';
import { TableData } from '../types/table-types';
import { ModifyTable } from '../modify/modify-table';

export default class ModifyTableHelper {
  static setTable = (data: TableData) => (
    element: XMLDocument | Element,
  ): void => {
    const modTable = new ModifyTable(element, data);
    modTable.modify()
    modTable.adjustWidth()
    modTable.adjustHeight();
  };

  static setTableData = (data: TableData) => (
    element: XMLDocument | Element,
  ): void => {
    const modTable = new ModifyTable(element, data);
    modTable.modify();
  };

  static adjustHeight = (data: TableData) => (
    element: XMLDocument | Element,
  ): void => {
    const modTable = new ModifyTable(element, data);
    modTable.adjustHeight();
  };

  static adjustWidth = (data: TableData) => (
    element: XMLDocument | Element,
  ): void => {
    const modTable = new ModifyTable(element, data);
    modTable.adjustWidth();
  };
}
