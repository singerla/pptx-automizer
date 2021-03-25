import {
  ChartData,
  ModificationPatternChildren, XYChartData,
} from '../types/types';
import { GeneralHelper } from './general-helper';
import { XmlHelper } from './xml-helper';
import StringIdGenerator from './string-id-generator';

export class ModifyChart {
  root: XMLDocument;
  StringIdGenerator: StringIdGenerator;
  data: XYChartData | ChartData;
  height: number;
  width: number;
  addCols: any[];
  addColsLength: number;

  constructor(root: XMLDocument, data: XYChartData | ChartData, addCols?: any[]) {
    this.root = root
    this.StringIdGenerator = new StringIdGenerator('ABCDEFGHIJKLMNOPQRSTUVWXYZ')

    this.data = data
    this.height = this.data.categories.length;
    this.addCols = GeneralHelper.arrayify(addCols);
    this.addColsLength = this.addCols.length;
    this.width = this.data.series.length + 1 + this.addColsLength;
  }

  pattern(
    pattern: ModificationPatternChildren,
    root?: XMLDocument | Element
  ): void {
    root = root || this.root

    for (const tag in pattern) {
      const parentPattern = pattern[tag];
      const index = parentPattern.index || 0;
      this.assert(root.getElementsByTagName(tag), index)
      const element = root.getElementsByTagName(tag)[index];

      if (GeneralHelper.propertyExists(parentPattern, 'modify')) {
        const modifies = GeneralHelper.arrayify(parentPattern.modify)
        Object.values(modifies).forEach(modify => modify(element))
      }

      if (GeneralHelper.propertyExists(parentPattern, 'children')) {
        this.pattern(parentPattern.children, element);
      }
    }
  }

  text = (label: string) => (element: Element): void => {
    element.firstChild.textContent = String(label);
  };

  value = (value: number | string, index?: number) => (element: Element): void => {
    element.getElementsByTagName('c:v')[0].firstChild.textContent = String(
      value,
    );
    if(index !== undefined) {
      element.setAttribute('idx', String(index));
    }
  };

  attribute = (attribute: string, value: string | number) => (
    element: Element,
  ): void => {
    element.setAttribute(attribute, String(value));
  };

  range = (series: number, length?: number) => (element: Element): void => {
    this.setRange(element, series, length);
  };

  setRange(element: Element, colId: number, length?: number): void {
    const range = element.firstChild.textContent;
    const info = range.split('!');
    const spans = info[1].split(':');
    const start = spans[0].split('$');
    const startRow = Number(spans[0].split('$')[2]);
    const colLetter = this.StringIdGenerator.start(colId).next()

    let endCell = ''
    if(length !== undefined) {
      const endRow = String(startRow + length - 1);
      endCell = `:$${colLetter}$${endRow}`;
    }

    const newRange = `${info[0]}!$${colLetter}$${start[2]}${endCell}`;

    element.firstChild.textContent = newRange;
  }

  getSpanString(startColNumber: number, startRowNumber:number, cols:number, rows:number): string {
    const startColLetter = this.StringIdGenerator.start(startColNumber).next()
    const endColLetter = this.StringIdGenerator.start(startColNumber+cols).next()
    const endRowNumber = startRowNumber + rows
    return `${startColLetter}${startRowNumber}:${endColLetter}${endRowNumber}`
  }

  getCellAddressString(c:number,r:number): string {
    const colLetter = this.StringIdGenerator.start(c).next()
    return `${colLetter}${r+1}`
  }

  assert(collection: HTMLCollectionOf<Element>, index: number) {
    if (!collection[index]) {
      const tplNode = collection[collection.length - 1];
      const newChild = tplNode.cloneNode(true);
      XmlHelper.insertAfter(newChild, tplNode);
    }
  }
}
