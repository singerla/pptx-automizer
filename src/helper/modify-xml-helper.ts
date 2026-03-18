import { Modification, ModificationTags } from '../types/modify-types';
import StringIdGenerator from './cell-id-helper';
import { GeneralHelper } from './general-helper';
import { XmlHelper } from './xml-helper';
import XmlElements from './xml-elements';
import { XmlDocument, XmlElement } from '../types/xml-types';

export default class ModifyXmlHelper {
  root: XmlDocument | XmlElement;
  templates: { [key: string]: XmlElement };

  constructor(root: XmlDocument | XmlElement) {
    this.root = root;
    this.templates = {};
  }

  modify(tags: ModificationTags, root?: XmlDocument | XmlElement): void {
    root = root || this.root;

    for (const tag in tags) {
      const modifier = tags[tag] as Modification;

      if (modifier.all) {
        this.modifyAll(tag, modifier, root);
      }

      if (modifier.collection) {
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
      } else {
        if (modifier.modify) {
          const modifies = GeneralHelper.arrayify(modifier.modify);
          Object.values(modifies).forEach((modifyXml) =>
            modifyXml(element as XmlElement),
          );
        }

        if (modifier.children) {
          this.modify(modifier.children, element as XmlElement);
        }
      }
    }
  }

  modifyAll(
    tag: string,
    modifier: Modification,
    root: XmlDocument | XmlElement,
  ): void {
    const elements = Array.from(root.getElementsByTagName(tag));
    elements.forEach((element) => {
      this.modify(modifier.children, element as XmlElement);
    });
  }

  assertElement(
    collection: HTMLCollectionOf<Element>,
    index: number,
    tag: string,
    parent: XmlDocument | XmlElement,
    modifier: Modification,
  ): XmlDocument | XmlElement | boolean {
    if (!collection[index]) {
      if (collection[collection.length - 1] === undefined) {
        this.createElement(parent, tag);
      } else {
        const lastSibling = collection[collection.length - 1];

        let sourceSibling = lastSibling;
        if (modifier.fromIndex !== undefined && modifier.fromIndex !== null && collection.item(modifier.fromIndex)) {
          // Store a clean template from the source element before it gets
          // modified, so that subsequent clones start from the original state.
          // Include a parent identifier in the key so that each parent context
          // (e.g. each c:dLbls within different c:ser) gets its own template.
          const parentId = (parent as XmlElement).tagName || 'root';
          const parentIndex = this.getParentIndex(parent as XmlElement);
          const fromIndexKey = parentId + '[' + parentIndex + ']:' + tag + ':fromIndex:' + modifier.fromIndex;
          if (!this.templates[fromIndexKey]) {
            this.templates[fromIndexKey] = collection.item(modifier.fromIndex).cloneNode(true) as XmlElement;
          }
          sourceSibling = this.templates[fromIndexKey];
        } else if (modifier.fromPrevious && collection.item(index - 1)) {
          sourceSibling = collection.item(index - 1);
        }

        if ((!sourceSibling || modifier.forceCreate) && this.templates[tag]) {
          sourceSibling = this.templates[tag];
        }

        const newChild = sourceSibling.cloneNode(true) as XmlElement;

        XmlHelper.insertAfter(newChild, lastSibling);
      }
    }

    const element = parent.getElementsByTagName(tag)[index];

    if (element) {
      this.templates[tag] =
        this.templates[tag] || (element.cloneNode(true) as XmlElement);
      return element;
    }

    return false;
  }

  getParentIndex(element: XmlElement): number {
    if (!element.parentNode) return 0;
    const siblings = (element.parentNode as XmlElement).getElementsByTagName(element.tagName);
    for (let i = 0; i < siblings.length; i++) {
      if (siblings[i] === element) return i;
    }
    return 0;
  }

  createElement(parent: XmlDocument | XmlElement, tag: string): boolean {
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
      case 'c:dLbls':
        new XmlElements(parent).dataPointLabels();
        return true;
      case 'c:dLbl':
        new XmlElements(parent).dataPointLabel();
        return true;
      case 'a:lnL':
      case 'a:lnR':
      case 'a:lnT':
      case 'a:lnB':
        new XmlElements(parent).tableCellBorder(tag);
        return true;
    }
    return false;
  }

  static getText = (element: XmlElement): string => {
    return element.firstChild.textContent;
  };

  static value =
    (value: number | string, index?: number) =>
    (element: XmlElement): void => {
      const valueElement = element.getElementsByTagName('c:v');
      if (!valueElement.length) {
        XmlHelper.dump(element);
        throw 'Unable to set value @index: ' + index;
      }

      if(!valueElement[0].firstChild) {
        return
      }

      valueElement[0].firstChild.textContent = XmlHelper.sanitizeText(value);
      if (index !== undefined) {
        element.setAttribute('idx', String(index));
      }
    };

  static textContent =
    (value: number | string) =>
    (element: XmlElement): void => {
      element.firstChild.textContent = XmlHelper.sanitizeText(value);
    };
  static attribute =
    (attribute: string, value: string | number) =>
    (element: XmlElement): void => {
      if (value != undefined)
        element.setAttribute(attribute, XmlHelper.sanitizeAttr(value));
    };

  static booleanAttribute =
    (attribute: string, state: boolean) =>
    (element: XmlElement): void => {
      element.setAttribute(attribute, state === true ? '1' : '0');
    };

  static range =
    (series: number, length?: number) =>
    (element: XmlElement): void => {
      const range = element.firstChild.textContent;
      element.firstChild.textContent = XmlHelper.sanitizeText(
        StringIdGenerator.setRange(
        range,
        series,
        length,
      ));
    };
}
