import { XmlHelper } from './xml-helper';
import { TableData } from '../types/table-types';

export default class ModifyTableHelper {
  /**
   * @TODO: Set table data of modify table helper
   */
  static setTableData = (data: TableData) => (
    element: XMLDocument | Document | Element,
  ): void => {
    console.log(data);
    XmlHelper.dump(element);
  };
}
