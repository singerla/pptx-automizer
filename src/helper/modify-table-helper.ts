import { ModifyTableParams, TableData, TableInfo } from '../types/table-types';
import { ModifyTable } from '../modify/modify-table';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { ShapeModificationCallback } from '../types/types';

export default class ModifyTableHelper {
  static setTable =
    (data: TableData, params?: ModifyTableParams) =>
    (element: XmlElement): void => {
      const modTable = new ModifyTable(element, data);

      if (params?.expand) {
        params?.expand.forEach((expand) => {
          const tableInfo = ModifyTableHelper.getTableInfo(element);
          const targetCell = tableInfo.find(
            (infoCell) => infoCell.textContent === expand.tag,
          );
          if (targetCell) {
            if (expand.mode === 'row') {
              modTable.expandRows(expand.count, targetCell.row);
            } else {
              if (targetCell.gridSpan) {
                modTable.expandSpanColumns(
                  expand.count,
                  targetCell.column,
                  targetCell.gridSpan,
                );
              } else {
                modTable.expandColumns(expand.count, targetCell.column);
              }
            }
          }
        });
      }

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
    (info?: TableInfo[]): ShapeModificationCallback =>
    (element: XmlElement): void => {
      if (Array.isArray(info)) {
        info.push(...ModifyTableHelper.getTableInfo(element));
      }
    };

  static getTableInfo = (element: XmlElement) => {
    const info = <TableInfo[]>[];
    const rows = element.getElementsByTagName('a:tr');
    if (!rows) {
      console.error("Can't find a table row.");
      return info;
    }

    for (let r = 0; r < rows.length; r++) {
      const row = rows.item(r);
      const columns = row.getElementsByTagName('a:tc');
      for (let c = 0; c < columns.length; c++) {
        const cell = columns.item(c);
        const gridSpan = cell.getAttribute('gridSpan');
        const hMerge = cell.getAttribute('hMerge');
        const texts = cell.getElementsByTagName('a:t');
        const text: string[] = [];
        for (let t = 0; t < texts.length; t++) {
          text.push(texts.item(t).textContent);
        }
        info.push({
          row: r,
          column: c,
          rowXml: row,
          columnXml: cell,
          text: text,
          textContent: text.join(''),
          gridSpan: Number(gridSpan),
          hMerge: Number(hMerge),
        });
      }
    }
    return info;
  };
}
