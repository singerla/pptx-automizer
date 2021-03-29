import {
  ChartData,
  ChartColumn,
  ModificationTags,
} from '../types/chart-types';
import { GeneralHelper } from '../helper/general-helper';
import { XmlHelper } from '../helper/xml-helper';
import StringIdGenerator from '../helper/cell-id-helper';

export class ModifyChart {
  root: XMLDocument;
  data: ChartData;
  height: number;
  width: number;
  columns: ChartColumn[];

  constructor(root: XMLDocument, data: ChartData, columns: ChartColumn[]) {
    this.root = root;
    this.data = data;
    this.columns = GeneralHelper.arrayify(columns);
    this.height = this.data.categories.length;
    this.width = this.columns.length;
  }

  pattern(
    tags: ModificationTags,
    root?: XMLDocument | Element,
  ): void {
    root = root || this.root;

    for (const tag in tags) {
      const parentPattern = tags[tag];
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
    const range = element.firstChild.textContent
    element.firstChild.textContent = StringIdGenerator.setRange(range, series, length);;
  };

  assertNode(collection: HTMLCollectionOf<Element>, index: number): void {
    if (!collection[index]) {
      const tplNode = collection[collection.length - 1];
      const newChild = tplNode.cloneNode(true);
      XmlHelper.insertAfter(newChild, tplNode);
    }
  }
}
