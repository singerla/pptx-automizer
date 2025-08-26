import { Color, TextStyle } from '../types/modify-types';
import ModifyColorHelper from './modify-color-helper';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlElement } from '../types/xml-types';
import { vd } from './general-helper';
import XmlElements from './xml-elements';
import { XmlHelper } from './xml-helper';
import {MultiTextParagraph} from '../interfaces/imulti-text';

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
  (element: XmlElement): void => {
    // Find or create the txBody element
    let txBody = element.getElementsByTagName('p:txBody')[0];
    if (!txBody) {
      txBody = element.ownerDocument.createElement('p:txBody');
      element.appendChild(txBody);

      // Create required bodyPr and lstStyle elements
      const bodyPr = element.ownerDocument.createElement('a:bodyPr');
      txBody.appendChild(bodyPr);

      const lstStyle = element.ownerDocument.createElement('a:lstStyle');
      txBody.appendChild(lstStyle);
    }

    // Clear existing paragraphs
    const existingParagraphs = txBody.getElementsByTagName('a:p');
    const length = existingParagraphs.length;
    for (let i = length - 1; i >= 0; i--) {
      const paragraph = existingParagraphs[i];
      paragraph.parentNode.removeChild(paragraph);
    }

    // Process each paragraph
    paragraphs.forEach(para => {
      // Create a new paragraph element
      const p = element.ownerDocument.createElement('a:p');
      txBody.appendChild(p);

      // Create paragraph properties
      const pPr = element.ownerDocument.createElement('a:pPr');
      p.appendChild(pPr);

      // Apply paragraph styling
      if (para.paragraph) {
        // Set bullet level
        if (para.paragraph.level !== undefined) {
          pPr.setAttribute('lvl', String(para.paragraph.level));
        }

        // Set bullet configuration
        if (para.paragraph.bullet) {
          const buChar = element.ownerDocument.createElement('a:buChar');
          buChar.setAttribute('char', '•'); // Default bullet character
          pPr.appendChild(buChar);
        } else if (para.paragraph.level === 0 || para.paragraph.bullet === false) {
          const buNone = element.ownerDocument.createElement('a:buNone');
          pPr.appendChild(buNone);
        }

        // Set alignment
        if (para.paragraph.alignment) {
          const algn = element.ownerDocument.createElement('a:algn');
          algn.setAttribute('val', para.paragraph.alignment);
          pPr.appendChild(algn);
        }

        // Set custom indentation
        if (para.paragraph.indent !== undefined) {
          pPr.setAttribute('indent', String(para.paragraph.indent));
        }

        // Set left margin
        if (para.paragraph.marginLeft !== undefined) {
          pPr.setAttribute('marL', String(para.paragraph.marginLeft));
        }

        // Set line spacing
        if (para.paragraph.lineSpacing !== undefined) {
          const lnSpc = element.ownerDocument.createElement('a:lnSpc');
          const spcPts = element.ownerDocument.createElement('a:spcPts');
          spcPts.setAttribute('val', String(Math.round(para.paragraph.lineSpacing * 100))); // Convert to 100ths of a point
          lnSpc.appendChild(spcPts);
          pPr.appendChild(lnSpc);
        }

        // Set space before paragraph
        if (para.paragraph.spaceBefore !== undefined) {
          const spcBef = element.ownerDocument.createElement('a:spcBef');
          const spcPts = element.ownerDocument.createElement('a:spcPts');
          spcPts.setAttribute('val', String(Math.round(para.paragraph.spaceBefore * 100))); // Convert to 100ths of a point
          spcBef.appendChild(spcPts);
          pPr.appendChild(spcBef);
        }

        // Set space after paragraph
        if (para.paragraph.spaceAfter !== undefined) {
          const spcAft = element.ownerDocument.createElement('a:spcAft');
          const spcPts = element.ownerDocument.createElement('a:spcPts');
          spcPts.setAttribute('val', String(Math.round(para.paragraph.spaceAfter * 100))); // Convert to 100ths of a point
          spcAft.appendChild(spcPts);
          pPr.appendChild(spcAft);
        }
      }

      // Handle text runs - if textRuns array exists, use it; otherwise, create a single run with paragraph text
      if (para.textRuns && para.textRuns.length > 0) {
        // Process each text run in the paragraph
        para.textRuns.forEach(run => {
          // Create text run
          const r = element.ownerDocument.createElement('a:r');
          p.appendChild(r);

          // Create text properties element
          const rPr = element.ownerDocument.createElement('a:rPr');
          r.appendChild(rPr);

          // Apply text styling if specified
          if (run.style) {
            ModifyTextHelper.style(run.style)(rPr);
          }

          // Create text element
          const t = element.ownerDocument.createElement('a:t');
          r.appendChild(t);

          // Handle empty strings with xml:space="preserve"
          if (run.text === '') {
            t.setAttribute('xml:space', 'preserve');
          }

          // Set text content
          const textNode = element.ownerDocument.createTextNode(run.text || '');
          t.appendChild(textNode);
        });
      } else if (para.text !== undefined) {
        // Create a single text run for the paragraph text
        const r = element.ownerDocument.createElement('a:r');
        p.appendChild(r);

        // Create text properties element
        const rPr = element.ownerDocument.createElement('a:rPr');
        r.appendChild(rPr);

        // Apply text styling
        if (para.style) {
          ModifyTextHelper.style(para.style)(rPr);
        }

        // Create text element
        const t = element.ownerDocument.createElement('a:t');
        r.appendChild(t);

        // Handle empty strings with xml:space="preserve"
        if (para.text === '') {
          t.setAttribute('xml:space', 'preserve');
        }

        // Set text content
        const textNode = element.ownerDocument.createTextNode(String(para.text || ''));
        t.appendChild(textNode);
      }
    });
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

  static htmlToText = (html: string) => {
    html =
      '<p><span style="font-size: 24px;">Testing layouts and exporting them.</span></p>\n' +
      '<ul>\n' +
      '<li>level 1 - 1</li>\n' +
      '<li>level 1 - 2</li>\n' +
      '<ul>\n' +
      '<li>level 1-2-1 <em>italics</em></li>\n' +
      '</ul>\n' +
      '<li>level 1 - 3</li>\n' +
      '<ul>\n' +
      '<li>level 1 - 3 - 1</li>\n' +
      '</ul>\n' +
      '</ul>\n' +
      '<p>Testing testing testing</p>\n' +
      '<p><strong>bold text</strong></p>\n';

    return (element: XmlElement): void => {
      XmlHelper.dump(element);
    };
  };
}
