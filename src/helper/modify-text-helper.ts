import { BulletListContent, Color, TextStyle } from '../types/modify-types';
import ModifyColorHelper from './modify-color-helper';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlElement } from '../types/xml-types';
import XmlElements from './xml-elements';
import { MultiTextParagraph } from '../interfaces/imulti-text';
import { MultiTextHelper } from './multitext-helper';
import { HtmlToMultiTextHelper } from './html-to-multitext-helper';
import { XmlHelper } from './xml-helper';

/**
 * Children of <a:rPr>/<a:defRPr>, in OOXML schema order
 * (CT_TextCharacterProperties). The style setters below are called from several
 * independent code paths (fill, highlight, typeface, hyperlink), so the final
 * child set is reordered instead of relying on insertion order - a wrong
 * sequence makes PowerPoint ask to repair the file.
 */
const RPR_CHILD_ORDER = [
  'a:ln',
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
  'a:effectLst',
  'a:effectDag',
  'a:highlight',
  'a:uLnTx',
  'a:uLn',
  'a:uFillTx',
  'a:uFill',
  'a:latin',
  'a:ea',
  'a:cs',
  'a:sym',
  'a:hlinkClick',
  'a:hlinkMouseOver',
  'a:rtl',
  'a:extLst',
];

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
    (list: BulletListContent) =>
    (element: XmlElement): void => {
      const xmlElements = new XmlElements(element);
      xmlElements.addBulletList(list);
    };

  static content =
    (label: number | string | undefined) =>
    (element: XmlElement): void => {
      if (label !== undefined && element.firstChild) {
        element.firstChild.textContent = XmlHelper.sanitizeText(label);
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
      if (style.isStrike !== undefined) {
        ModifyTextHelper.setStrike(style.isStrike)(element);
      }
      if (style.fontFamily !== undefined) {
        ModifyTextHelper.setFontFamily(style.fontFamily)(element);
      }
      if (style.highlight !== undefined) {
        ModifyTextHelper.setHighlight(style.highlight)(element);
      }

      // The setters above append; the schema wants a fixed sequence
      XmlHelper.sortChildrenBySchema(element, RPR_CHILD_ORDER);
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
   * Set strikethrough attribute on text
   */
  static setStrike =
    (isStrike: boolean) =>
    (element: XmlElement): void => {
      ModifyXmlHelper.attribute(
        'strike',
        isStrike ? 'sngStrike' : 'noStrike',
      )(element);
    };

  /**
   * Set the typeface of text inside an <a:rPr> element.
   * Only the latin script slot is set - `a:ea`/`a:cs` stay inherited.
   */
  static setFontFamily =
    (fontFamily: string) =>
    (element: XmlElement): void => {
      if (!fontFamily) return;

      let latin = XmlHelper.getFirstDirectChild(element, ['a:latin']);
      if (!latin) {
        latin = element.ownerDocument.createElement('a:latin');
        XmlHelper.insertInSchemaOrder(element, latin, RPR_CHILD_ORDER);
      }

      latin.setAttribute('typeface', XmlHelper.sanitizeAttr(fontFamily));
    };

  /**
   * Set the highlight (marker pen) color of text inside an <a:rPr> element
   */
  static setHighlight =
    (color: Color) =>
    (element: XmlElement): void => {
      if (!color || !color.type) return;

      ModifyColorHelper.normalizeColorObject(color);

      let highlight = XmlHelper.getFirstDirectChild(element, ['a:highlight']);
      if (!highlight) {
        highlight = element.ownerDocument.createElement('a:highlight');
        XmlHelper.insertInSchemaOrder(element, highlight, RPR_CHILD_ORDER);
      }

      XmlHelper.removeAllChildren(highlight);
      highlight.appendChild(new XmlElements(element, { color }).colorType());
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
