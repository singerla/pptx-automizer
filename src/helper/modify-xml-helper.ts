import { Modification, ModificationTags } from '../types/modify-types';

import StringIdGenerator from './cell-id-helper';
import { GeneralHelper, vd } from './general-helper';
import { XmlHelper } from './xml-helper';
import XmlElements, { XmlElementParams } from './xml-elements';

export default class ModifyXmlHelper {
  root: XMLDocument | Element;
  templates: { [key: string]: Node };

  constructor(root: XMLDocument | Element) {
    this.root = root;
    this.templates = {};
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
      const isRequired =
        modifier.isRequired !== undefined ? modifier.isRequired : true;

      const element = this.assertElement(
        root.getElementsByTagName(tag),
        index,
        tag,
        root,
        modifier,
      );

      if (element === false) {
        if (isRequired === true) {
          // vd('Could not assert required tag: ' + tag + '@index:' + index);
        }
        return;
      }

      if (GeneralHelper.propertyExists(modifier, 'modify')) {
        const modifies = GeneralHelper.arrayify(modifier.modify);
        Object.values(modifies).forEach((modifyXml) =>
          modifyXml(element as Element),
        );
      }

      if (GeneralHelper.propertyExists(modifier, 'children')) {
        this.modify(modifier.children, element as Element);
      }
    }
  }

  assertElement(
    collection: HTMLCollectionOf<Element>,
    index: number,
    tag: string,
    parent: XMLDocument | Element,
    modifier: Modification,
  ): XMLDocument | Element | boolean {
    if (!collection[index]) {
      if (collection[collection.length - 1] === undefined) {
        this.createElement(parent, tag);
      } else {
        const previousSibling = collection[collection.length - 1];

        const newChild =
          this.templates[tag] && !modifier.fromPrevious
            ? this.templates[tag].cloneNode(true)
            : previousSibling.cloneNode(true);

        XmlHelper.insertAfter(newChild, previousSibling);
      }
    }

    const element = parent.getElementsByTagName(tag)[index];

    if (element) {
      this.templates[tag] = this.templates[tag] || element.cloneNode(true);
      return element;
    }

    return false;
  }

  createElement(parent: XMLDocument | Element, tag: string): boolean {
    switch (tag) {
      case 'a:t':
        new XmlElements(parent).text();
        return true;
      case 'c:dPt':
        new XmlElements(parent).dataPoint();
        return true;
      case 'c:spPr':
        new XmlElements(parent).shapeProperties();
        return true;
      case 'c:dLbl':
        new XmlElements(parent).dataPointLabel();
        return true;
    }
    return false;
  }

  static getText = (element: Element): string => {
    return element.firstChild.textContent;
  };

  static value =
    (value: number | string, index?: number) =>
    (element: Element): void => {
      element.getElementsByTagName('c:v')[0].firstChild.textContent =
        String(value);
      if (index !== undefined) {
        element.setAttribute('idx', String(index));
      }
    };

  static attribute =
    (attribute: string, value: string | number) =>
    (element: Element): void => {
      element.setAttribute(attribute, String(value));
    };

  static booleanAttribute =
    (attribute: string, state: boolean) =>
    (element: Element): void => {
      element.setAttribute(attribute, state === true ? '1' : '0');
    };

  static range =
    (series: number, length?: number) =>
    (element: Element): void => {
      const range = element.firstChild.textContent;
      element.firstChild.textContent = StringIdGenerator.setRange(
        range,
        series,
        length,
      );
    };
}
