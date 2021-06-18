
import { XmlHelper } from '../helper/xml-helper';
import ModifyXmlHelper from '../helper/modify-xml-helper';
import { TableData, TableRow } from '../types/table-types';
import { Modification, ModificationTags } from '../types/modify-types';

export class ModifyTable {
  data: TableData;
  table: ModifyXmlHelper;
  xml: XMLDocument | Element;

  constructor(
    table: XMLDocument | Element,
    data: TableData,
  ) {
    this.data = data;

    this.table = new ModifyXmlHelper(table);
    this.xml = table;
  }

  modify(): ModifyTable {
    this.setContents()
    this.sliceRows()

    return this
  }

  setContents() {
    this.data.body.forEach((row: TableRow, r: number) => {
      row.values.forEach((cell: number | string, c: number) => {
        this.table.modify(
          this.row(r, 
            this.column(c, 
              this.cell(cell)
            )
          )
        )
      })
    })
  }

  sliceRows() {
    this.table.modify({
      'a:tbl': this.slice('a:tr', this.data.body.length)
    })
  }

  row = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'a:tr': {
        index: index,
        children: children
      }
    }
  }

  column = (index: number, children: ModificationTags): ModificationTags => {
    return {
      'a:tc': {
        index: index,
        children: children
      }
    }
  }

  cell = (value: number | string): ModificationTags => {
    return {
      'a:t': {
        modify: ModifyXmlHelper.text(value),
      }
    }
  }

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
    const tableHeight = Number(this.xml
      .getElementsByTagName('p:xfrm')[0]
      .getElementsByTagName('a:ext')[0]
      .getAttribute('cy'))

    const rowHeight = tableHeight / this.data.body.length

    this.data.body.forEach((row: TableRow, r: number) => {
      this.table.modify({
        'a:tr': {
          index: r,
          modify: ModifyXmlHelper.attribute('h', Math.round(rowHeight)),
        }
      })
    })
  }
}
