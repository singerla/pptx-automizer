import { Color, TextStyle } from '../types/modify-types';
import ModifyColorHelper from './modify-color-helper';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlElement } from '../types/xml-types';
import XmlElements from './xml-elements';
import { MultiTextParagraph } from '../interfaces/imulti-text';
import { MultiTextHelper } from './multitext-helper';
import { HtmlToMultiTextHelper } from './html-to-multitext-helper';
import { XmlHelper } from './xml-helper';
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

  static setMultiText =
    (paragraphs: MultiTextParagraph[]) =>
    (element: XmlElement, relation?: XmlElement): void => {
      new MultiTextHelper(element, relation).run(paragraphs);
    };

  static htmlToMultiText = (html: string) => {
    const paragraphs = new HtmlToMultiTextHelper().run(html);
    return (element: XmlElement, relation?: XmlElement): void => {
      this.setMultiText(paragraphs)(element, relation);
    };
  };

  static setBulletList =
    (list) =>
    (element: XmlElement): void => {
      const xmlElements = new XmlElements(element);
      xmlElements.addBulletList(list);
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
      if (style.isUnderlined !== undefined) {
        ModifyTextHelper.setUnderlined(style.isUnderlined)(element);
      }
      if (style.isSuperscript !== undefined) {
        ModifyTextHelper.setSuperscript(style.isSuperscript)(element);
      }
      if (style.isSubscript !== undefined) {
        ModifyTextHelper.setSubscript(style.isSubscript)(element);
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

  /**
   * Set underlined attribute on text
   */
  static setUnderlined =
    (isUnderlined: boolean) =>
    (element: XmlElement): void => {
      if (isUnderlined) {
        element.setAttribute('u', 'sng');
      }
    };

  /**
   * Set superscript attribute on text
   */
  static setSuperscript =
    (isSuperscript: boolean) =>
      (element: XmlElement): void => {
        ModifyXmlHelper.attribute('baseline', isSuperscript ? '30000' : '0')(element);
      };

  /**
   * Set subscript attribute on text
   */
  static setSubscript =
    (isSubscript: boolean) =>
      (element: XmlElement): void => {
        ModifyXmlHelper.attribute('baseline', isSubscript ? '-25000' : '0')(element);
      };

  /**
   * Set bullet type (font and character) for bullet points
   */
  static setBulletType =
    (font: string, character: string) =>
    (element: XmlElement): void => {
      const paragraphs = element.getElementsByTagName('a:p');
      XmlHelper.modifyCollection(paragraphs, (paragraph) => {
        const pPr = paragraph.getElementsByTagName('a:pPr')[0];

        if (!pPr) {
          return;
        }

        const existingBuFont = pPr.getElementsByTagName('a:buFont')[0];
        if (existingBuFont) {
          existingBuFont.setAttribute('typeface', font);
        }

        const existingBuChar = pPr.getElementsByTagName('a:buChar')[0];
        if (existingBuChar) {
          existingBuChar.setAttribute('char', character);
        }

        const existingBuBlip = pPr.getElementsByTagName('a:buBlip')[0];
        if (existingBuBlip) {
          XmlHelper.remove(existingBuBlip)
        }
      });
    };
}
