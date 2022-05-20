import { XmlHelper } from './xml-helper';
import { TableData, ModifyTableParams } from '../types/table-types';
import { ModifyTable } from '../modify/modify-table';

export default class ModifyTableHelper {
  static setTable =
    (data: TableData, params?: ModifyTableParams) =>
    (element: XMLDocument | Element): void => {
      const modTable = new ModifyTable(element, data);

      modTable.modify();

      if (params?.setHeight) {
        modTable.setSize('cy', params.setHeight);
      }
      if (params?.setWidth) {
        modTable.setSize('cx', params.setWidth);
      }
      if (!params || params?.adjustHeight) {
        modTable.adjustHeight();
      }
      if (!params || params?.adjustWidth) {
        modTable.adjustWidth();
      }
    };

  static setTableData =
    (data: TableData) =>
    (element: XMLDocument | Element): void => {
      const modTable = new ModifyTable(element, data);
      modTable.modify();
    };

  static adjustHeight =
    (data: TableData) =>
    (element: XMLDocument | Element): void => {
      const modTable = new ModifyTable(element, data);
      modTable.adjustHeight();
    };

  static adjustWidth =
    (data: TableData) =>
    (element: XMLDocument | Element): void => {
      const modTable = new ModifyTable(element, data);
      modTable.adjustWidth();
    };
}
