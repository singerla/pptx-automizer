import { ModifyTableParams, TableData } from '../types/table-types';
import { ModifyTable } from '../modify/modify-table';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { XmlSlideHelper } from './xml-slide-helper';

export default class ModifyTableHelper {
  static setTable =
    (data: TableData, params?: ModifyTableParams) =>
    (element: XmlElement): void => {
      const modTable = new ModifyTable(element, data);

      if (params?.expand) {
        params?.expand.forEach((expand) => {
          const tableInfo = XmlSlideHelper.readTableInfo(element);
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

      modTable.modify(params);

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

  static updateColumnWidth =
    (index: number, size: number) =>
    (element: XmlDocument | XmlElement): void => {
      const modTable = new ModifyTable(element);
      modTable.updateColumnWidth(index, size);
    };
  static updateRowHeight =
    (index: number, size: number) =>
    (element: XmlDocument | XmlElement): void => {
      const modTable = new ModifyTable(element);
      modTable.updateRowHeight(index, size);
    };

  static setTableStyle =
    (styleId: string, attribs: string[]) => (element: XmlElement) => {
      const tblPr = element.getElementsByTagName('a:tblPr').item(0);

      const setTableStyleId = (tableStyleId: XmlElement, id: string) => {
        tableStyleId.textContent = id;
      };

      const createTableStyleId = (tblPr: XmlElement) => {
        const tableStyleId =
          tblPr.ownerDocument.createElement('a:tableStyleId');
        tblPr.appendChild(tableStyleId);
        return tableStyleId;
      };
      const updateTable = (tblPr: XmlElement) => {
        [
          'firstRow',
          'firstCol',
          'lastRow',
          'lastCol',
          'bandRow',
          'bandCol',
        ].forEach((attrib) => {
          if (attribs.includes(attrib)) {
            tblPr.setAttribute(attrib, '1');
          } else {
            tblPr.removeAttribute(attrib);
          }
        });

        const tableStyleId =
          element.getElementsByTagName('a:tableStyleId').item(0) ||
          createTableStyleId(tblPr);

        if (tableStyleId) {
          setTableStyleId(tableStyleId, styleId);
        }
      };
      if (tblPr) {
        updateTable(tblPr);
      }
    };
}
