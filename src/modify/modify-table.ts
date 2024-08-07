import { XmlHelper } from '../helper/xml-helper';
import ModifyXmlHelper from '../helper/modify-xml-helper';
import { TableData, TableRow, TableRowStyle } from '../types/table-types';
import { Border, Modification, ModificationTags } from '../types/modify-types';
import ModifyTextHelper from '../helper/modify-text-helper';
import { ModifyColorHelper } from '../index';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { GeneralHelper, vd } from '../helper/general-helper';

export class ModifyTable {
  data: TableData;
  table: ModifyXmlHelper;
  xml: XmlDocument | XmlElement;

  constructor(table: XmlDocument | XmlElement, data?: TableData) {
    this.data = data;

    this.table = new ModifyXmlHelper(table);
    this.xml = table;
  }

  modify(): ModifyTable {
    this.setRows();
    this.setGridCols();

    this.sliceRows();
    this.sliceCols();

    return this;
  }

  setRows() {
    this.data.body.forEach((row: TableRow, r: number) => {
      row.values.forEach((cell: number | string, c: number) => {
        const rowStyles = row.styles && row.styles[c] ? row.styles[c] : {};
        this.table.modify(
          this.row(r, this.column(c, this.cell(cell, rowStyles))),
        );
        this.table.modify({
          'a16:rowId': {
            index: r,
            modify: ModifyXmlHelper.attribute('val', r),
          },
        });
      });
    });
  }

  setGridCols() {
    this.data.body[0]?.values.forEach((cell, c: number) => {
      this.table.modify({
        'a:gridCol': {
          index: c,
        },
        'a16:colId': {
          index: c,
          modify: ModifyXmlHelper.attribute('val', c),
        },
      });
    });
  }

  sliceRows() {
    this.table.modify({
      'a:tbl': this.slice('a:tr', this.data.body.length),
    });
  }

  sliceCols() {
    this.table.modify({
      'a:tblGrid': this.slice('a:gridCol', this.data.body[0]?.values.length),
    });
  }

  row = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'a:tr': {
        index: index,
        children: children,
      },
    };
  };

  column = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'a:tc': {
        index: index,
        children: children,
      },
    };
  };

  cell = (value: number | string, style?: TableRowStyle): ModificationTags => {
    return {
      'a:t': {
        modify: ModifyTextHelper.content(value),
      },
      'a:rPr': {
        modify: ModifyTextHelper.style(style),
      },
      'a:tcPr': {
        ...this.setCellStyle(style),
      },
    };
  };

  setCellStyle(style) {
    const cellProps = {
      modify: [],
      children: {},
    };

    if (style.background) {
      cellProps.modify.push(
        ModifyColorHelper.solidFill(style.background, 'last'),
      );
    }

    if (style.border) {
      cellProps.children = this.setCellBorder(style);
    }

    return cellProps;
  }

  setCellBorder(style) {
    const borders = GeneralHelper.arrayify<Border>(style.border);
    const sortBorderTags = ['lnB', 'lnT', 'lnR', 'lnL'];
    const modifications = {};
    borders
      .sort((b1, b2) =>
        sortBorderTags.indexOf(b1.tag) < sortBorderTags.indexOf(b2.tag)
          ? -1
          : 1,
      )
      .forEach((border) => {
        const tag = 'a:' + border.tag;

        const modifyCell = [];

        if (border.color) {
          modifyCell.push(ModifyColorHelper.solidFill(border.color));
        }
        if (border.weight) {
          modifyCell.push(ModifyXmlHelper.attribute('w', border.weight));
        }

        modifications[tag] = {
          modify: modifyCell,
        };

        if (border.type) {
          modifications[tag].children = {
            'a:prstDash': {
              modify: ModifyXmlHelper.attribute('val', border.type),
            },
          };
        }
      });

    return modifications;
  }

  slice(tag: string, length: number): Modification {
    return {
      children: {
        [tag]: {
          collection: (collection: HTMLCollectionOf<XmlElement>) => {
            XmlHelper.sliceCollection(collection, length);
          },
        },
      },
    };
  }

  adjustHeight() {
    const tableHeight = this.getTableSize('cy');
    const rowHeight = tableHeight / this.data.body.length;

    this.data.body.forEach((row: TableRow, r: number) => {
      this.table.modify({
        'a:tr': {
          index: r,
          modify: ModifyXmlHelper.attribute('h', Math.round(rowHeight)),
        },
      });
    });

    return this;
  }

  adjustWidth() {
    const tableWidth = this.getTableSize('cx');
    const rowWidth = tableWidth / this.data.body[0]?.values?.length || 1;

    this.data.body[0]?.values.forEach((cell, c: number) => {
      this.table.modify({
        'a:gridCol': {
          index: c,
          modify: ModifyXmlHelper.attribute('w', Math.round(rowWidth)),
        },
      });
    });

    return this;
  }

  setSize(orientation: 'cx' | 'cy', size: number): void {
    const sizeElement = this.xml
      .getElementsByTagName('p:xfrm')[0]
      .getElementsByTagName('a:ext')[0];

    sizeElement.setAttribute(orientation, String(size));
  }

  getTableSize(orientation: string): number {
    return Number(
      this.xml
        .getElementsByTagName('p:xfrm')[0]
        .getElementsByTagName('a:ext')[0]
        .getAttribute(orientation),
    );
  }

  expandRows = (count: number, rowId: number) => {
    const tplRow = this.xml.getElementsByTagName('a:tr').item(rowId);
    for (let r = 1; r <= count; r++) {
      const newRow = tplRow.cloneNode(true) as XmlElement;
      XmlHelper.insertAfter(newRow, tplRow);
      this.updateId(newRow, 'a16:rowId', r);
    }
  };

  expandSpanColumns = (count: number, colId: number, gridSpan: number) => {
    for (let cs = 1; cs <= count; cs++) {
      const rows = this.xml.getElementsByTagName('a:tr');
      for (let r = 0; r < rows.length; r++) {
        const row = rows.item(r);
        const columns = row.getElementsByTagName('a:tc');
        const maxC = colId + gridSpan;
        for (let c = colId; c < maxC; c++) {
          const sourceCell = columns.item(c);
          const insertAfter = columns.item(c + gridSpan - 1);
          const clone = sourceCell.cloneNode(true) as XmlElement;
          XmlHelper.insertAfter(clone, insertAfter);
        }
      }
    }
    this.expandGrid(count, colId, gridSpan);
  };

  expandColumns = (count: number, colId: number) => {
    const rows = this.xml.getElementsByTagName('a:tr');
    for (let r = 0; r < rows.length; r++) {
      const row = rows.item(r);
      const columns = row.getElementsByTagName('a:tc');
      const sourceCell = columns.item(colId);
      const newCell = this.getExpandCellClone(columns, sourceCell, colId);

      XmlHelper.insertAfter(newCell, sourceCell);
    }

    this.expandGrid(count, colId, 1);
  };

  getExpandCellClone(
    columns: HTMLCollectionOf<XmlElement>,
    sourceCell: XmlElement,
    colId: number,
  ): XmlElement {
    const gridSpan = sourceCell.getAttribute('gridSpan');
    const hMerge = sourceCell.getAttribute('hMerge');

    if (gridSpan) {
      const incrementGridSpan = Number(gridSpan) + 1;
      sourceCell.setAttribute('gridSpan', String(incrementGridSpan));
      return columns.item(colId + 1).cloneNode(true) as XmlElement;
    }

    if (hMerge) {
      for (let findCol = colId - 1; colId >= 0; colId--) {
        const previousSibling = columns.item(findCol);
        if (!previousSibling) {
          break;
        }
        const hasSpan = previousSibling.getAttribute('gridSpan');
        if (hasSpan) {
          const incrementGridSpan = Number(hasSpan) + 1;
          previousSibling.setAttribute('gridSpan', String(incrementGridSpan));
          break;
        }
      }
    }

    return sourceCell.cloneNode(true) as XmlElement;
  }

  expandGrid = (count: number, colId: number, gridSpan: number) => {
    const tblGrid = this.xml.getElementsByTagName('a:tblGrid').item(0);
    for (let cs = 1; cs <= count; cs++) {
      const maxC = colId + gridSpan;
      for (let c = colId; c < maxC; c++) {
        const sourceTblGridCol = tblGrid
          .getElementsByTagName('a:gridCol')
          .item(c);
        const newCol = sourceTblGridCol.cloneNode(true) as XmlElement;
        XmlHelper.insertAfter(newCol, sourceTblGridCol);
        this.updateId(newCol, 'a16:colId', c * (cs + 1) * colId * 1000);
      }
    }
  };

  updateId = (element: XmlElement, tag: string, id: number) => {
    const idElement = element.getElementsByTagName(tag).item(0);
    const previousId = Number(idElement.getAttribute('val'));
    idElement.setAttribute('val', String(previousId + id));
  };
}
