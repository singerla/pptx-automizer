import { XmlHelper } from '../helper/xml-helper';
import ModifyXmlHelper from '../helper/modify-xml-helper';
import { TableData, TableRow } from '../types/table-types';
import {Color, Modification, ModificationTags, TextStyle} from '../types/modify-types';
import { vd } from '../helper/general-helper';
import ModifyTextHelper from '../helper/modify-text-helper';

export class ModifyTable {
  data: TableData;
  table: ModifyXmlHelper;
  xml: XMLDocument | Element;

  constructor(table: XMLDocument | Element, data?: TableData) {
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
        const rowStyles = (row.styles && row.styles[c]) ? row.styles[c] : {}
        this.table.modify(
          this.row(r,
            this.column(c,
              this.cell(cell, rowStyles)
            )
          )
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
    this.data.body[0].values.forEach((cell, c: number) => {
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
      'a:tblGrid': this.slice('a:gridCol', this.data.body[0].values.length),
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

  cell = (value: number | string, style?: TextStyle): ModificationTags => {
    return {
      'a:t': {
        modify: ModifyTextHelper.content(value),
      },
      'a:rPr': {
        modify: ModifyTextHelper.style(style),
      },
    };
  };

  slice(tag: string, length: number): Modification {
    return {
      children: {
        [tag]: {
          collection: (collection: HTMLCollectionOf<Element>) => {
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
    const rowWidth = tableWidth / this.data.body[0].values.length;

    this.data.body[0].values.forEach((cell, c: number) => {
      this.table.modify({
        'a:gridCol': {
          index: c,
          modify: ModifyXmlHelper.attribute('w', Math.round(rowWidth)),
        },
      });
    });

    return this;
  }

  getTableSize(orientation: string): number {
    return Number(
      this.xml
        .getElementsByTagName('p:xfrm')[0]
        .getElementsByTagName('a:ext')[0]
        .getAttribute(orientation),
    );
  }
}
