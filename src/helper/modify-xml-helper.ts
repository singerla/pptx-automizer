import { ElementNotFoundError } from '../errors';
import { Modification, ModificationTags } from '../types/modify-types';
import StringIdGenerator from './cell-id-helper';
import { GeneralHelper } from './general-helper';
import { XmlHelper } from './xml-helper';
import XmlElements, {
  C_DLBLS_CHILD_ORDER,
  C_SER_CHILD_ORDER,
} from './xml-elements';
import { XmlDocument, XmlElement, XmlElementCollection } from '../types/xml-types';
import { log } from './logger';

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

      const element =
        modifier.matchIdx !== undefined
          ? this.assertElementByIdx(tag, root, modifier)
          : this.assertElement(
              root.getElementsByTagName(tag),
              index,
              tag,
              root,
              modifier,
            );

      if (element === false) {
        const target =
          modifier.matchIdx !== undefined
            ? tag + '@c:idx:' + modifier.matchIdx
            : tag + '@index:' + index;
        if (isRequired === true) {
          log.warn('Could not assert required tag: ' + target);
        } else {
          log.debug('Skipped modification of absent optional tag: ' + target);
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
    collection: XmlElementCollection,
    index: number,
    tag: string,
    parent: XmlDocument | XmlElement,
    modifier: Modification,
  ): XmlDocument | XmlElement | boolean {
    if (!collection[index]) {
      if (modifier.isRequired === false) {
        // "Modify if present, never create": absence is a regular no-op.
        return false;
      }

      if (collection[collection.length - 1] === undefined) {
        this.createElement(parent, tag);
      } else {
        const lastSibling = collection[collection.length - 1];

        let sourceSibling = lastSibling;
        const template = this.getFromIndexTemplate(
          collection,
          tag,
          parent,
          modifier,
        );
        if (template) {
          sourceSibling = template;
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

  /**
   * Resolve the target among a tag's elements by the value of their
   * `<c:idx val="…"/>` child instead of by sibling position — the correct
   * addressing for sparse chart collections (`c:dPt`, `c:dLbl`), where one
   * element exists per *explicitly styled* point and `c:idx` names the
   * category. See `Modification.matchIdx`.
   *
   * A missing element is created — cloned from the clean `fromIndex`
   * template when given, built as a minimal shell otherwise — stamped with
   * the requested idx and inserted so ascending `c:idx` order is kept.
   */
  assertElementByIdx(
    tag: string,
    parent: XmlDocument | XmlElement,
    modifier: Modification,
  ): XmlElement | false {
    const matchIdx = modifier.matchIdx;
    const collection = parent.getElementsByTagName(tag);

    for (let i = 0; i < collection.length; i++) {
      const element = collection.item(i) as XmlElement;
      if (ModifyXmlHelper.getIdxValue(element) === matchIdx) {
        return element;
      }
    }

    if (modifier.isRequired === false) {
      return false;
    }

    const template = this.getFromIndexTemplate(
      collection,
      tag,
      parent,
      modifier,
    );
    const newElement = template
      ? (template.cloneNode(true) as XmlElement)
      : this.buildElement(parent, tag);

    if (!newElement) {
      return false;
    }

    const idx = newElement.getElementsByTagName('c:idx').item(0);
    if (!idx) {
      return false;
    }
    idx.setAttribute('val', String(matchIdx));

    const successor = Array.from(collection).find(
      (sibling) =>
        ModifyXmlHelper.getIdxValue(sibling as XmlElement) > matchIdx,
    ) as XmlElement | undefined;

    if (successor) {
      successor.parentNode.insertBefore(newElement, successor);
    } else if (collection.length > 0) {
      XmlHelper.insertAfter(
        newElement,
        collection.item(collection.length - 1) as XmlElement,
      );
    } else {
      const order =
        tag === 'c:dLbl' ? C_DLBLS_CHILD_ORDER : C_SER_CHILD_ORDER;
      XmlHelper.insertInSchemaOrder(parent as XmlElement, newElement, order);
    }

    return newElement;
  }

  static getIdxValue(element: XmlElement): number {
    const idx = element.getElementsByTagName('c:idx').item(0);
    return idx ? Number(idx.getAttribute('val')) : NaN;
  }

  /**
   * A clean clone of `collection[fromIndex]`, taken before that element got
   * modified, so subsequent clones start from the original state. Cached per
   * parent context: each c:dLbls within a different c:ser gets its own
   * template.
   */
  getFromIndexTemplate(
    collection: XmlElementCollection,
    tag: string,
    parent: XmlDocument | XmlElement,
    modifier: Modification,
  ): XmlElement | null {
    if (
      modifier.fromIndex === undefined ||
      modifier.fromIndex === null ||
      !collection.item(modifier.fromIndex)
    ) {
      return null;
    }
    const parentId = (parent as XmlElement).tagName || 'root';
    const parentIndex = this.getParentIndex(parent as XmlElement);
    const fromIndexKey =
      parentId + '[' + parentIndex + ']:' + tag + ':fromIndex:' + modifier.fromIndex;
    if (!this.templates[fromIndexKey]) {
      this.templates[fromIndexKey] = collection
        .item(modifier.fromIndex)
        .cloneNode(true) as XmlElement;
    }
    return this.templates[fromIndexKey];
  }

  /**
   * Index of `element` among all elements of its tag name in the whole
   * document — a stable identifier for the template cache key, so equally
   * named parents in different subtrees (e.g. the c:dLbls of each c:ser)
   * do not collide.
   */
  getParentIndex(element: XmlElement): number {
    const scope = element.ownerDocument || element.parentNode;
    if (!scope) return 0;
    const siblings = (scope as XmlElement).getElementsByTagName(
      element.tagName,
    );
    for (let i = 0; i < siblings.length; i++) {
      if (siblings[i] === element) return i;
    }
    return 0;
  }

  /**
   * Build a detached minimal element for `tag`, or null when the tag is not
   * supported. Unlike `createElement`, insertion is left to the caller.
   */
  buildElement(parent: XmlDocument | XmlElement, tag: string): XmlElement | null {
    switch (tag) {
      case 'c:dPt':
        return new XmlElements(parent).buildDataPoint();
      case 'c:dLbl':
        return new XmlElements(parent).buildDataPointLabel();
    }
    return null;
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
      case 'a:ln':
        new XmlElements(parent).plainLine();
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
        throw new ElementNotFoundError('Unable to set value @index: ' + index, {
          selector: 'c:v',
        });
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

  static removeAttribute =
    (attribute: string) =>
    (element: XmlElement): void => {
      element.removeAttribute(attribute);
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
