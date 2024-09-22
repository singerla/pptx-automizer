import { CreateTreeAttributes, CreateTreeStructure, Modification, ModificationTags } from '../types/modify-types';
import StringIdGenerator from './cell-id-helper';
import { GeneralHelper } from './general-helper';
import { XmlHelper } from './xml-helper';
import XmlElements from './xml-elements';
import { XmlDocument, XmlElement } from '../types/xml-types';

export default class ModifyXmlHelper {
  root: XmlDocument | XmlElement;
  treeStart: XmlDocument | XmlElement;
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
        if (modifier.fromIndex && collection.item(modifier.fromIndex)) {
          sourceSibling = collection.item(modifier.fromIndex);
        } else if (modifier.fromPrevious && collection.item(index - 1)) {
          sourceSibling = collection.item(index - 1);
        }

        if ((!sourceSibling || modifier.forceCreate) && this.templates[tag]) {
          sourceSibling = this.templates[tag];
        }

        const newChild = sourceSibling.cloneNode(true);

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

  createTree(structure: CreateTreeStructure, root?: XmlElement): XmlElement {
    root = root || this.root as XmlElement;
    this.treeStart = root;
    this.buildTree(structure, root);
    this.reorderChildren(root, structure);

    return root;
  }

  private buildTree(structure: CreateTreeStructure, currentParent: XmlElement): XmlElement {
    for (const tag in structure) {

      if (structure[tag].remove) {
        const elements = currentParent.getElementsByTagName(tag);
        for (let i = elements.length - 1; i >= 0; i--) {
          currentParent.removeChild(elements[i]);
        }
        continue;
      }

      let element: XmlElement;
      if (structure[tag].unique) {
        element = this.ensureUniqueElement(currentParent, tag);
      } else if (structure[tag].clone) {
        element = this.cloneOrBuildElement(tag, currentParent, structure[tag].attributes);
      } else {
        element = currentParent.getElementsByTagName(tag)[0] || this.buildElement(tag, structure[tag].attributes);
        if (!currentParent.getElementsByTagName(tag)[0]) {
          currentParent.appendChild(element);
        }
      }

      if (structure[tag].attributes) {
        this.addAttributes(element, structure[tag].attributes);
      }
      if (structure[tag].removeChildren) {
        while (element.firstChild) {
          element.removeChild(element.firstChild);
        }
      }

      if (structure[tag].empty) {
        continue;
      }

      if (structure[tag].slot !== undefined) {
        element.textContent = this.setTextElement(String(structure[tag].slot));
      }

      if (structure[tag].children) {
        this.buildTree(structure[tag].children, element);
      }
    }
    return currentParent;
  }

  private reorderChildren(parent: XmlElement, structure: CreateTreeStructure): void {
    for (const tag in structure) {
      const elements = Array.from(parent.getElementsByTagName(tag));
      if (elements.length > 0) {
        const element = elements[0];

        const childTags = Object.keys(structure[tag].children || {});

        childTags.forEach(childTag => {
          const children = Array.from(element.getElementsByTagName(childTag));
          children.forEach(child => element.appendChild(child));
        });

        this.reorderChildren(element, structure[tag].children || {});
      }
    }
  }

  private cloneOrBuildElement(tag: string, currentParent: XmlElement, attributes?: CreateTreeAttributes): XmlElement {
    let element = currentParent.getElementsByTagName(tag)[0];
    if (!element) {
      element = this.cloneElement(tag) || this.buildElement(tag, attributes);
      currentParent.appendChild(element);
    }
    return element;
  }

  private cloneElement(tag: string): XmlElement | null {
    const elementToClone = this.treeStart.getElementsByTagName(tag)[0];
    if (elementToClone) {
      return elementToClone.cloneNode(true) as XmlElement;
    }
    return null;
  }

  private ensureUniqueElement(parent: XmlElement, tag: string): XmlElement {
    let element = parent.getElementsByTagName(tag)[0];
    if (!element) {
      element = this.buildElement(tag);
      parent.appendChild(element);
    } else {
      this.removeAdditionalElements(parent, tag, 0);
    }
    return element;
  }

  addAttributes(element: XmlElement, attributes: CreateTreeAttributes): void {
    for (const [attr, value] of Object.entries(attributes)) {
      if (typeof value === 'string' || typeof value === 'number') {
        element.setAttribute(attr, String(value));
      }
    }
  }

  setTextElement(text: string): string {
    if (typeof text === 'string' || typeof text === 'number') {
      if (text.length > 0) {
        return text;
      }
    }
    return ' ';
  }

  removeAdditionalElements(parent: XmlElement, tag: string, startIndex: number): void {
    const elements = parent.getElementsByTagName(tag);
    for (let i = elements.length - 1; i > startIndex; i--) {
      parent.removeChild(elements[i]);
    }
  }

  private buildElement(tag: string, attributes?: CreateTreeAttributes): XmlElement {
    const element = this.root.ownerDocument.createElement(tag);
    if (attributes) {
      for (const [attr, value] of Object.entries(attributes)) {
        if (typeof value === 'string' || typeof value === 'number') {
          element.setAttribute(attr, String(value));
        }
      }
    }
    return element;
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

      valueElement[0].firstChild.textContent = String(value);
      if (index !== undefined) {
        element.setAttribute('idx', String(index));
      }
    };

  static textContent =
    (value: number | string) =>
    (element: XmlElement): void => {
      element.firstChild.textContent = String(value);
    };
  static attribute =
    (attribute: string, value: string | number) =>
    (element: XmlElement): void => {
      if (value != undefined) element.setAttribute(attribute, String(value));
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
      element.firstChild.textContent = StringIdGenerator.setRange(
        range,
        series,
        length,
      );
    };
}
