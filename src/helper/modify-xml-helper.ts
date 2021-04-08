import { Modification, ModificationTags } from '../types/modify-types';

import StringIdGenerator from './cell-id-helper';
import { GeneralHelper } from './general-helper';
import { XmlHelper } from './xml-helper';

export default class ModifyXmlHelper {
  root: XMLDocument;

  constructor(root: XMLDocument) {
    this.root = root;
  }

  modify(tags: ModificationTags, root?: XMLDocument | Element): void {
    root = root || this.root;

    for (const tag in tags) {
      const modifier = tags[tag] as Modification;

      if (GeneralHelper.propertyExists(modifier, 'collection')) {
        const modifies = GeneralHelper.arrayify(modifier.collection);
        const collection = root.getElementsByTagName(tag);
        Object.values(modifies).forEach((modifyXml) => modifyXml(collection));
        return;
      }

      const index = modifier.index || 0;

      this.assertNode(root.getElementsByTagName(tag), index, tag, modifier);
      const element = root.getElementsByTagName(tag)[index];

      if (GeneralHelper.propertyExists(modifier, 'modify')) {
        const modifies = GeneralHelper.arrayify(modifier.modify);
        Object.values(modifies).forEach((modifyXml) => modifyXml(element));
      }

      if (GeneralHelper.propertyExists(modifier, 'children')) {
        this.modify(modifier.children, element);
      }
    }
  }

  static text = (label: string) => (element: Element): void => {
    element.firstChild.textContent = String(label);
  };

  static value = (value: number | string, index?: number) => (
    element: Element,
  ): void => {
    element.getElementsByTagName('c:v')[0].firstChild.textContent = String(
      value,
    );
    if (index !== undefined) {
      element.setAttribute('idx', String(index));
    }
  };

  static attribute = (attribute: string, value: string | number) => (
    element: Element,
  ): void => {
    element.setAttribute(attribute, String(value));
  };

  static range = (series: number, length?: number) => (
    element: Element,
  ): void => {
    const range = element.firstChild.textContent;
    element.firstChild.textContent = StringIdGenerator.setRange(
      range,
      series,
      length,
    );
  };

  assertNode(
    collection: HTMLCollectionOf<Element>,
    index: number,
    tag?: string,
    info?,
  ): void {
    if (!collection[index]) {
      if (collection[collection.length - 1] === undefined) {
        console.log(info);
        throw new Error(`Index ${index} not found at "${tag}"`);
      }
      const tplNode = collection[collection.length - 1];
      const newChild = tplNode.cloneNode(true);
      XmlHelper.insertAfter(newChild, tplNode);
    }
  }
}
