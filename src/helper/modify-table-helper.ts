import { XmlHelper } from './xml-helper';
import { TableData } from '../types/table-types';
import { ModifyTable } from '../modify/modify-table';

export default class ModifyTableHelper {
  /**
   * @TODO: Set table data of modify table helper
   */
  static setTableData = (data: TableData) => (
    element: XMLDocument | Element,
  ): void => {
    const modTable = new ModifyTable(element, data);

    modTable.modify().adjustHeight();

    // console.log(data);
    // XmlHelper.dump(element);
  };
}
