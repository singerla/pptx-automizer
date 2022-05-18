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
      const isRequired =
        modifier.isRequired !== undefined ? modifier.isRequired : true;
      const forceCreate =
        modifier.forceCreate !== undefined ? modifier.forceCreate : false;

      if (tag === 'c:dLbl') {
        // XmlHelper.dump(root.parentNode as any);
      }

      const asserted = this.assertNode(
        root.getElementsByTagName(tag),
        index,
        tag,
        root,
        isRequired,
        forceCreate,
      );

      if (asserted === false) {
        if (isRequired === true) {
          // XmlHelper.dump(root)
          // vd(tags)
          vd('Could not assert required tag: ' + tag + '@index:' + index);
        }
        return;
      }

      const element = root.getElementsByTagName(tag)[index];

      if (tag === 'c:dLbl') {
        // console.log(root.getElementsByTagName(tag).length);
        // XmlHelper.dump(element.parentNode.parentNode as any);
      }

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

  assertNode(
    collection: HTMLCollectionOf<Element>,
    index: number,
    tag: string,
    parent: XMLDocument | Element,
    required: boolean,
    forceCreate: boolean,
  ): boolean {
    if (forceCreate) {
      return this.createNode(parent, tag, index, required);
    }

    if (!collection[index]) {
      if (collection[collection.length - 1] === undefined) {
        return this.createNode(parent, tag, index, required);
      } else {
        const tplNode = collection[collection.length - 1];
        const newChild = tplNode.cloneNode(true);
        XmlHelper.insertAfter(newChild, tplNode);
      }
    }
    return true;
  }

  createNode(
    parent: XMLDocument | Element,
    tag: string,
    index: number,
    required: boolean,
  ): boolean {
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
    //
    // if(required === true) {
    //   throw new Error(
    //     `Could not create new node at index ${index}; Tag "${tag}"`,
    //   );
    // }

    return false;
  }
}
