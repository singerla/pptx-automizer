import {
  ChartData,
  ChartColumn,
  ModificationPatternChildren,
} from '../types/chart-types';
import { GeneralHelper } from '../helper/general-helper';
import { XmlHelper } from '../helper/xml-helper';
import StringIdGenerator from '../helper/string-id-generator';

export class ModifyChart {
  root: XMLDocument;
  StringIdGenerator: StringIdGenerator;
  data: ChartData;
  height: number;
  width: number;
  columns: ChartColumn[];

  constructor(root: XMLDocument, data: ChartData, columns: ChartColumn[]) {
    this.root = root;
    this.StringIdGenerator = new StringIdGenerator(
      'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
    );
    this.data = data;
    this.columns = GeneralHelper.arrayify(columns);
    this.height = this.data.categories.length;
    this.width = this.columns.length;
  }

  pattern(
    pattern: ModificationPatternChildren,
    root?: XMLDocument | Element,
  ): void {
    root = root || this.root;

    for (const tag in pattern) {
      const parentPattern = pattern[tag];
      const index = parentPattern.index || 0;
      this.assertNode(root.getElementsByTagName(tag), index);
      const element = root.getElementsByTagName(tag)[index];

      if (GeneralHelper.propertyExists(parentPattern, 'modify')) {
        const modifies = GeneralHelper.arrayify(parentPattern.modify);
        Object.values(modifies).forEach((modify) => modify(element));
      }

      if (GeneralHelper.propertyExists(parentPattern, 'children')) {
        this.pattern(parentPattern.children, element);
      }
    }
  }

  text = (label: string) => (element: Element): void => {
    element.firstChild.textContent = String(label);
  };

  value = (value: number | string, index?: number) => (
    element: Element,
  ): void => {
    element.getElementsByTagName('c:v')[0].firstChild.textContent = String(
      value,
    );
    if (index !== undefined) {
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
    const colLetter = this.StringIdGenerator.start(colId).next();

    let endCell = '';
    if (length !== undefined) {
      const endRow = String(startRow + length - 1);
      endCell = `:$${colLetter}$${endRow}`;
    }

    const newRange = `${info[0]}!$${colLetter}$${start[2]}${endCell}`;

    element.firstChild.textContent = newRange;
  }

  getSpanString(
    startColNumber: number,
    startRowNumber: number,
    cols: number,
    rows: number,
  ): string {
    const startColLetter = this.StringIdGenerator.start(startColNumber).next();
    const endColLetter = this.StringIdGenerator.start(
      startColNumber + cols,
    ).next();
    const endRowNumber = startRowNumber + rows;
    return `${startColLetter}${startRowNumber}:${endColLetter}${endRowNumber}`;
  }

  getCellAddressString(c: number, r: number): string {
    const colLetter = this.StringIdGenerator.start(c).next();
    return `${colLetter}${r + 1}`;
  }

  assertNode(collection: HTMLCollectionOf<Element>, index: number): void {
    if (!collection[index]) {
      const tplNode = collection[collection.length - 1];
      const newChild = tplNode.cloneNode(true);
      XmlHelper.insertAfter(newChild, tplNode);
    }
  }
}
