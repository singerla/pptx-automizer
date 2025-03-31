import { XmlHelper } from '../helper/xml-helper';
import ModifyXmlHelper from '../helper/modify-xml-helper';
import {
  ModifyTableParams,
  TableData,
  TableRow,
  TableRowStyle,
} from '../types/table-types';
import {
  Border,
  Modification,
  ModificationTags,
  ModifyCallback,
} from '../types/modify-types';
import ModifyTextHelper from '../helper/modify-text-helper';
import { ModifyColorHelper } from '../index';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { GeneralHelper, vd } from '../helper/general-helper';

export class ModifyTable {
  data: TableData;
  table: ModifyXmlHelper;
  xml: XmlDocument | XmlElement;
  maxCols = 0;
  params: ModifyTableParams;

  constructor(table: XmlDocument | XmlElement, data?: TableData) {
    this.data = data;

    this.table = new ModifyXmlHelper(table);
    this.xml = table;

    this.data?.body.forEach((row) => {
      this.maxCols =
        row.values.length > this.maxCols ? row.values.length : this.maxCols;
    });
  }

  modify(params?: ModifyTableParams): ModifyTable {
    this.params = params;

    this.setRows();
    this.setGridCols();

    this.sliceRows();
    this.sliceCols();

    return this;
  }

  setRows() {
    const alreadyExpanded = this.params?.expand?.find(
      (expand) => expand.mode === 'column',
    );

    this.data.body.forEach((row: TableRow, r: number) => {
      row.values.forEach((cell, c: number) => {
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
        if (this.params?.expand && !alreadyExpanded) {
          this.expandOtherMergedCellsInColumn(c, r);
        }
      });
    });
  }

  expandOtherMergedCellsInColumn(c: number, r: number) {
    const rows = this.xml.getElementsByTagName('a:tr');
    for (let rs = 0; rs < rows.length; rs++) {
      // Skip current row
      if (r !== rs) {
        const row = rows.item(r);
        const columns = row.getElementsByTagName('a:tc');
        const sourceCell = columns.item(c);
        this.expandGridSpan(sourceCell);
      }
    }
  }

  setGridCols() {
    for (let c = 0; c <= this.maxCols; c++) {
      this.table.modify({
        'a:gridCol': {
          index: c,
        },
        'a16:colId': {
          index: c,
          modify: ModifyXmlHelper.attribute('val', c),
        },
      });
    }
  }

  sliceRows() {
    this.table.modify({
      'a:tbl': this.slice('a:tr', this.data.body.length),
    });
  }

  sliceCols() {
    this.table.modify({
      'a:tblGrid': this.slice('a:gridCol', this.maxCols),
    });
  }

  row = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'a:tr': {
        forceCreate: true,
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
        fromPrevious: !!this.params?.expand,
      },
    };
  };

  cell = (value: number | string, style?: TableRowStyle): ModificationTags => {
    return {
      'a:txBody': {
        children: {
          'a:t': {
            modify: ModifyTextHelper.content(value),
          },
          'a:rPr': {
            modify: ModifyTextHelper.style(style),
          },
          'a:r': {
            collection: (collection: HTMLCollectionOf<Element>) => {
              XmlHelper.sliceCollection(collection, 1);
            },
          },
        },
      },
      'a:tcPr': {
        ...this.setCellStyle(style),
      },
    };
  };

  setCellStyle(style: TableRowStyle) {
    const cellProps: Modification & {
      modify: ModifyCallback[];
    } = {
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

  updateColumnWidth(c: number, size: number) {
    const tableWidth = this.getTableSize('cx');
    const targetSize = Math.round(size);
    let currentSize = 0;

    this.table.modify({
      'a:gridCol': {
        index: c,
        modify: [
          (ele) => {
            currentSize = Number(ele.getAttribute('w'));
          },
          ModifyXmlHelper.attribute('w', targetSize),
        ],
      },
    });

    const diff = currentSize - targetSize;
    const targetWidth = tableWidth - diff;

    this.setSize('cx', targetWidth);

    return this;
  }

  updateRowHeight(r: number, size: number) {
    const tableSize = this.getTableSize('cy');
    const targetSize = Math.round(size);
    let currentSize = 0;

    this.table.modify({
      'a:tr': {
        index: r,
        modify: [
          (ele) => {
            currentSize = Number(ele.getAttribute('h'));
          },
          ModifyXmlHelper.attribute('h', targetSize),
        ],
      },
    });

    const diff = currentSize - targetSize;
    const targetTableSize = tableSize - diff;

    this.setSize('cy', targetTableSize);

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
    for (let cs = 1; cs <= count; cs++) {
      const rows = this.xml.getElementsByTagName('a:tr');
      for (let r = 0; r < rows.length; r++) {
        const row = rows.item(r);
        const columns = row.getElementsByTagName('a:tc');
        const sourceCell = columns.item(colId);

        const newCell = this.getExpandCellClone(columns, sourceCell, colId);
        XmlHelper.insertAfter(newCell, sourceCell);
      }
    }

    this.expandGrid(count, colId, 1);
  };

  getExpandCellClone(
    columns: HTMLCollectionOf<XmlElement>,
    sourceCell: XmlElement,
    colId: number,
  ): XmlElement {
    const hasGridSpan = this.expandGridSpan(sourceCell);
    if (hasGridSpan) {
      return columns.item(colId + 1).cloneNode(true) as XmlElement;
    }

    const hMerge = sourceCell.getAttribute('hMerge');
    if (hMerge) {
      for (let findCol = colId - 1; colId >= 0; colId--) {
        const previousSibling = columns.item(findCol);
        if (!previousSibling) {
          break;
        }
        const siblingHasSpan = this.expandGridSpan(previousSibling);
        if (siblingHasSpan) {
          break;
        }
      }
    }

    return sourceCell.cloneNode(true) as XmlElement;
  }

  expandGridSpan(sourceCell: XmlElement) {
    const gridSpan = sourceCell.getAttribute('gridSpan');
    if (gridSpan) {
      const incrementGridSpan = Number(gridSpan) + 1;
      sourceCell.setAttribute('gridSpan', String(incrementGridSpan));
      return true;
    }
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
