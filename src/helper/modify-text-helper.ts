import { Color, TextStyle } from '../types/modify-types';
import ModifyColorHelper from './modify-color-helper';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlElement } from '../types/xml-types';
import { vd } from './general-helper';

export default class ModifyTextHelper {
  /**
   * Set text content of first paragraph and remove remaining block/paragraph elements.
   */
  static setText =
    (text: number | string) =>
    (element: XmlElement): void => {
      const paragraphs = element.getElementsByTagName('a:p');
      const length = paragraphs.length;
      for (let i = 0; i < length; i++) {
        const paragraph = paragraphs[i];
        if (i === 0) {
          const blocks = element.getElementsByTagName('a:r');
          const length = blocks.length;
          for (let j = 0; j < length; j++) {
            const block = blocks[j];
            if (j === 0) {
              const textNode = block.getElementsByTagName('a:t')[0];
              ModifyTextHelper.content(text)(textNode);
            } else {
              block.parentNode.removeChild(block);
            }
          }
        } else {
          paragraph.parentNode.removeChild(paragraph);
        }
      }
    };

  static setBulletList =
    (list) => (element: XmlElement): void => {
      const namespaceURIs = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
      };
      const doc = element.ownerDocument;

      let txBody = element.getElementsByTagName('p:txBody')[0];
      if (!txBody) {
        txBody = doc.createElementNS(namespaceURIs['p'], 'p:txBody');
        element.appendChild(txBody);
      } else {
        while (txBody.firstChild) {
          txBody.removeChild(txBody.firstChild);
        }
      }

      const bodyPr = doc.createElementNS(namespaceURIs['a'], 'a:bodyPr');
      txBody.appendChild(bodyPr);
      const lstStyle = doc.createElementNS(namespaceURIs['a'], 'a:lstStyle');
      txBody.appendChild(lstStyle);

      const processList = (items, level) => {
        items.forEach((item) => {
          if (Array.isArray(item)) {
            processList(item, level + 1);
          } else {
            const p = doc.createElementNS(namespaceURIs['a'], 'a:p');

            const pPr = doc.createElementNS(namespaceURIs['a'], 'a:pPr');
            if (level > 0) {
              pPr.setAttribute('lvl', String(level));
            }
            p.appendChild(pPr);

            const r = doc.createElementNS(namespaceURIs['a'], 'a:r');

            const rPr = doc.createElementNS(namespaceURIs['a'], 'a:rPr');
            r.appendChild(rPr);

            const t = doc.createElementNS(namespaceURIs['a'], 'a:t');
            const textNode = doc.createTextNode(String(item));
            t.appendChild(textNode);

            r.appendChild(t);
            p.appendChild(r);
            txBody.appendChild(p);
          }
        });
      };

      processList(list, 0);
    };

  static content =
    (label: number | string | undefined) =>
    (element: XmlElement): void => {
      if (label !== undefined && element.firstChild) {
        element.firstChild.textContent = String(label);
      }
    };

  /**
   * Set text style inside an <a:rPr> element
   */
  static style =
    (style: TextStyle) =>
    (element: XmlElement): void => {
      if (!style) return;
      if (style.color !== undefined) {
        ModifyTextHelper.setColor(style.color)(element);
      }
      if (style.size !== undefined) {
        ModifyTextHelper.setSize(style.size)(element);
      }
      if (style.isBold !== undefined) {
        ModifyTextHelper.setBold(style.isBold)(element);
      }
      if (style.isItalics !== undefined) {
        ModifyTextHelper.setItalics(style.isItalics)(element);
      }
    };

  /**
   * Set color of text insinde an <a:rPr> element
   */
  static setColor =
    (color: Color) =>
    (element: XmlElement): void => {
      ModifyColorHelper.solidFill(color)(element);
    };

  /**
   * Set size of text inside an <a:rPr> element
   */
  static setSize =
    (size: number) =>
    (element: XmlElement): void => {
      if (!size) return;
      element.setAttribute('sz', String(Math.round(size)));
    };

  /**
   * Set bold attribute on text
   */
  static setBold =
    (isBold: boolean) =>
    (element: XmlElement): void => {
      ModifyXmlHelper.booleanAttribute('b', isBold)(element);
    };

  /**
   * Set italics attribute on text
   */
  static setItalics =
    (isItalics: boolean) =>
    (element: XmlElement): void => {
      ModifyXmlHelper.booleanAttribute('i', isItalics)(element);
    };
}
