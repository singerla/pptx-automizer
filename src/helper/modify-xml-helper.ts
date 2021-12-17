import { Modification, ModificationTags } from '../types/modify-types';

import StringIdGenerator from './cell-id-helper';
import { GeneralHelper, vd } from './general-helper';
import { XmlHelper } from './xml-helper';
import XmlElements, { XmlElementParams } from './xml-elements';

export default class ModifyXmlHelper {
  root: XMLDocument | Element;

  constructor(root: XMLDocument | Element) {
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

      this.assertNode(root.getElementsByTagName(tag), index, tag, root);
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

  static getText = (element: Element): string => {
    return element.firstChild.textContent;
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

  static booleanAttribute = (attribute: string, state: boolean) => (element: Element): void => {
    element.setAttribute(attribute, (state === true) ? '1' : '0');
  }

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
    tag: string,
    parent: XMLDocument | Element,
  ): void {
    if (!collection[index]) {
      if (collection[collection.length - 1] === undefined) {
        this.createNode(parent, tag, index);
      } else {
        const tplNode = collection[collection.length - 1];
        const newChild = tplNode.cloneNode(true);
        XmlHelper.insertAfter(newChild, tplNode);
      }
    }
  }

  createNode(parent: XMLDocument | Element, tag: string, index: number): void {
    switch (tag) {
      case 'a:t':
        new XmlElements(parent).text();
        return;
    }

    throw new Error(
      `Could not create new node at index ${index}; Tag "${tag}"`,
    );
  }
}
