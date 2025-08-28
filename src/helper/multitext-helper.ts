import { Color, TextStyle } from '../types/modify-types';
import ModifyColorHelper from './modify-color-helper';
import ModifyXmlHelper from './modify-xml-helper';
import { XmlElement } from '../types/xml-types';
import { vd } from './general-helper';
import XmlElements from './xml-elements';
import { XmlHelper } from './xml-helper';
import { MultiTextParagraph } from '../interfaces/imulti-text';
import ModifyTextHelper from './modify-text-helper';

export class MultiTextHelper {
  private element: XmlElement;
  private document: Document;

  constructor(element: XmlElement) {
    this.element = element;
    this.document = element.ownerDocument;
  }

  /**
   * Apply multi-text paragraphs to the element
   */
  run(paragraphs: MultiTextParagraph[]): void {
    const txBody = this.getOrCreateTxBody();
    this.clearExistingParagraphs(txBody);
    this.createParagraphs(txBody, paragraphs);
  }

  /**
   * Find or create the txBody element and ensure it has required child elements
   */
  private getOrCreateTxBody(): XmlElement {
    let txBody = this.element.getElementsByTagName('p:txBody')[0];

    if (!txBody) {
      txBody = this.document.createElement('p:txBody');
      this.element.appendChild(txBody);

      // Create required bodyPr and lstStyle elements
      const bodyPr = this.document.createElement('a:bodyPr');
      txBody.appendChild(bodyPr);

      const lstStyle = this.document.createElement('a:lstStyle');
      txBody.appendChild(lstStyle);
    }

    return txBody;
  }

  /**
   * Remove all existing paragraph elements
   */
  private clearExistingParagraphs(txBody: XmlElement): void {
    const existingParagraphs = txBody.getElementsByTagName('a:p');
    const length = existingParagraphs.length;

    for (let i = length - 1; i >= 0; i--) {
      const paragraph = existingParagraphs[i];
      paragraph.parentNode.removeChild(paragraph);
    }
  }

  /**
   * Create paragraph elements for each MultiTextParagraph
   */
  private createParagraphs(txBody: XmlElement, paragraphs: MultiTextParagraph[]): void {
    paragraphs.forEach(para => {
      const p = this.document.createElement('a:p');
      txBody.appendChild(p);

      const pPr = this.document.createElement('a:pPr');
      p.appendChild(pPr);

      if (para.paragraph) {
        this.applyParagraphProperties(pPr, para.paragraph);
      }

      if (para.textRuns && para.textRuns.length > 0) {
        this.createTextRuns(p, para.textRuns);
      } else if (para.text !== undefined) {
        this.createSingleTextRun(p, para.text, para.style);
      }
    });
  }

  /**
   * Apply paragraph styling properties
   */
  private applyParagraphProperties(pPr: XmlElement, paragraphProps: any): void {
    // Set bullet level
    if (paragraphProps.level !== undefined) {
      pPr.setAttribute('lvl', String(paragraphProps.level));
    }

    // Set bullet configuration
    this.applyBulletConfiguration(pPr, paragraphProps);

    // Set alignment
    if (paragraphProps.alignment !== undefined) {
      pPr.setAttribute('algn', paragraphProps.alignment);
    }

    // Set custom indentation
    if (paragraphProps.indent !== undefined) {
      pPr.setAttribute('indent', String(paragraphProps.indent));
    }

    // Set left margin
    if (paragraphProps.marginLeft !== undefined) {
      pPr.setAttribute('marL', String(paragraphProps.marginLeft));
    }

    // Apply spacing properties
    this.applySpacingProperties(pPr, paragraphProps);
  }

  /**
   * Apply bullet configuration to paragraph properties
   */
  private applyBulletConfiguration(pPr: XmlElement, paragraphProps: any): void {
    if (paragraphProps.bullet) {
      const buChar = this.document.createElement('a:buChar');
      buChar.setAttribute('char', '•'); // Default bullet character
      pPr.appendChild(buChar);
    } else if (paragraphProps.level === 0 || paragraphProps.bullet === false) {
      const buNone = this.document.createElement('a:buNone');
      pPr.appendChild(buNone);
    }
  }

  /**
   * Apply spacing properties to paragraph properties
   */
  private applySpacingProperties(pPr: XmlElement, paragraphProps: any): void {
    // Set line spacing
    if (paragraphProps.lineSpacing !== undefined) {
      const lnSpc = this.document.createElement('a:lnSpc');
      const spcPts = this.document.createElement('a:spcPts');
      spcPts.setAttribute(
        'val',
        String(Math.round(paragraphProps.lineSpacing * 100))
      ); // Convert to 100ths of a point
      lnSpc.appendChild(spcPts);
      pPr.appendChild(lnSpc);
    }

    // Set space before paragraph
    if (paragraphProps.spaceBefore !== undefined) {
      const spcBef = this.document.createElement('a:spcBef');
      const spcPts = this.document.createElement('a:spcPts');
      spcPts.setAttribute(
        'val',
        String(Math.round(paragraphProps.spaceBefore * 100))
      ); // Convert to 100ths of a point
      spcBef.appendChild(spcPts);
      pPr.appendChild(spcBef);
    }

    // Set space after paragraph
    if (paragraphProps.spaceAfter !== undefined) {
      const spcAft = this.document.createElement('a:spcAft');
      const spcPts = this.document.createElement('a:spcPts');
      spcPts.setAttribute(
        'val',
        String(Math.round(paragraphProps.spaceAfter * 100))
      ); // Convert to 100ths of a point
      spcAft.appendChild(spcPts);
      pPr.appendChild(spcAft);
    }
  }

  /**
   * Create text runs for a paragraph
   */
  private createTextRuns(p: XmlElement, textRuns: Array<{ text: string; style?: TextStyle }>): void {
    textRuns.forEach(run => {
      const r = this.document.createElement('a:r');
      p.appendChild(r);

      const rPr = this.document.createElement('a:rPr');
      r.appendChild(rPr);

      // Apply text styling if specified
      if (run.style) {
        ModifyTextHelper.style(run.style)(rPr);
      }

      this.createTextElement(r, run.text || '');
    });
  }

  /**
   * Create a single text run with the given text and style
   */
  private createSingleTextRun(p: XmlElement, text: string | number, style?: TextStyle): void {
    const r = this.document.createElement('a:r');
    p.appendChild(r);

    const rPr = this.document.createElement('a:rPr');
    r.appendChild(rPr);

    // Apply text styling
    if (style) {
      ModifyTextHelper.style(style)(rPr);
    }

    this.createTextElement(r, String(text || ''));
  }

  /**
   * Create a text element with the given content
   */
  private createTextElement(r: XmlElement, text: string): void {
    const t = this.document.createElement('a:t');
    r.appendChild(t);

    // Handle empty strings with xml:space="preserve"
    if (text === '') {
      t.setAttribute('xml:space', 'preserve');
    }

    // Set text content
    const textNode = this.document.createTextNode(text);
    t.appendChild(textNode);
  }
}
