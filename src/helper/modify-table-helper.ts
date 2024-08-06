import { ModifyTableParams, TableData } from '../types/table-types';
import { ModifyTable } from '../modify/modify-table';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { ShapeModificationCallback } from '../types/types';

export default class ModifyTableHelper {
  static setTable =
    (data: TableData, params?: ModifyTableParams) =>
    (element: XmlDocument | XmlElement): void => {
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
    (element: XmlDocument | XmlElement): void => {
      const modTable = new ModifyTable(element, data);
      modTable.modify();
    };

  static adjustHeight =
    (data: TableData) =>
    (element: XmlDocument | XmlElement): void => {
      const modTable = new ModifyTable(element, data);
      modTable.adjustHeight();
    };

  static adjustWidth =
    (data: TableData) =>
    (element: XmlDocument | XmlElement): void => {
      const modTable = new ModifyTable(element, data);
      modTable.adjustWidth();
    };

  static readTableData =
    (info?: TableData): ShapeModificationCallback =>
    (element: XmlElement): void => {
      const body = [];
      const rows = element.getElementsByTagName('a:tr');
      for (let r = 0; r < rows.length; r++) {
        const row = rows.item(r);
        const columns = row.getElementsByTagName('a:tc');
        for (let c = 0; c < columns.length; c++) {
          const cell = columns.item(c);
          const gridSpan = cell.getAttribute('gridSpan');
          const texts = cell.getElementsByTagName('a:t');
          const text: string[] = [];
          for (let t = 0; t < texts.length; t++) {
            text.push(texts.item(t).textContent);
          }
          body.push({
            row: r,
            column: c,
            rowXml: row,
            columnXml: cell,
            text: text.join(' '),
            gridSpan: Number(gridSpan),
          });
        }
      }
      if (typeof info === 'object') {
        info.body = body;
      }
    };
}
