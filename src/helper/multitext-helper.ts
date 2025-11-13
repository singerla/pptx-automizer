import { TextStyle } from '../types/modify-types';
import { XmlElement } from '../types/xml-types';
import { MultiTextParagraph } from '../interfaces/imulti-text';
import ModifyTextHelper from './modify-text-helper';
import { XmlHelper } from './xml-helper';
import HyperlinkElement from './modify-hyperlink-element';
import { Logger } from './general-helper';

export class MultiTextHelper {
  private element: XmlElement;
  private document: Document;
  private relationElement?: XmlElement;

  constructor(element: XmlElement, relationElement?: XmlElement) {
    this.element = element;
    this.document = element.ownerDocument;
    this.relationElement = relationElement;
  }

  /**
   * Apply multi-text paragraphs to the element
   */
  run(paragraphs: MultiTextParagraph[]): void {
    const txBody = this.getOrCreateTxBody();
    const defaultStyle = this.extractDefaultStyle(txBody);
    this.clearExistingParagraphs(txBody);
    this.createParagraphs(txBody, paragraphs, defaultStyle);
  }

  /**
   * Extract default style from existing paragraphs
   */
  private extractDefaultStyle(txBody: XmlElement): TextStyle {
    const defaultStyle: TextStyle = {};
    const existingParagraphs = txBody.getElementsByTagName('a:p');

    if (existingParagraphs.length === 0) {
      return defaultStyle;
    }

    // Try to get font size from the first text run
    const firstPara = existingParagraphs[0];
    const firstRun = firstPara.getElementsByTagName('a:r')[0];

    if (firstRun) {
      const rPr = firstRun.getElementsByTagName('a:rPr')[0];
      if (rPr) {
        // Extract font size if it exists
        const fontSize = rPr.getAttribute('sz');
        if (fontSize) {
          defaultStyle.size = parseInt(fontSize);
        }

        // Extract color if it exists
        const solidFill = rPr.getElementsByTagName('a:solidFill')[0];
        if (solidFill) {
          const srgbClr = solidFill.getElementsByTagName('a:srgbClr')[0];
          if (srgbClr) {
            const colorValue = srgbClr.getAttribute('val');
            if (colorValue) {
              defaultStyle.color = {
                type: 'srgbClr',
                value: colorValue,
              };
            }
          }
        }

        // Extract bold and italic if they exist
        const bold = rPr.getAttribute('b');
        if (bold === '1') {
          defaultStyle.isBold = true;
        }

        const italic = rPr.getAttribute('i');
        if (italic === '1') {
          defaultStyle.isItalics = true;
        }
      }
    }

    return defaultStyle;
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
  private createParagraphs(
    txBody: XmlElement,
    paragraphs: MultiTextParagraph[],
    defaultStyle: TextStyle = {},
  ): void {
    paragraphs.forEach((para) => {
      const p = this.document.createElement('a:p');
      txBody.appendChild(p);

      const pPr = this.document.createElement('a:pPr');
      p.appendChild(pPr);

      if (para.paragraph) {
        this.applyParagraphProperties(pPr, para.paragraph);
      }

      if (para.textRuns && para.textRuns.length > 0) {
        this.createTextRuns(p, para.textRuns, defaultStyle);
      } else if (para.text !== undefined) {
        // Merge default style with provided style
        const mergedStyle = this.mergeStyles(defaultStyle, para.style);
        this.createSingleTextRun(p, para.text, mergedStyle);
      }
    });
  }

  /**
   * Merge default style with provided style
   */
  private mergeStyles(
    defaultStyle: TextStyle = {},
    customStyle: TextStyle = {},
  ): TextStyle {
    return {
      ...defaultStyle,
      ...customStyle,
    };
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
        String(Math.round(paragraphProps.lineSpacing * 100)),
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
        String(Math.round(paragraphProps.spaceBefore * 100)),
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
        String(Math.round(paragraphProps.spaceAfter * 100)),
      ); // Convert to 100ths of a point
      spcAft.appendChild(spcPts);
      pPr.appendChild(spcAft);
    }
  }

  /**
   * Create text runs for a paragraph
   */
  private createTextRuns(
    p: XmlElement,
    textRuns: Array<{ text: string; style?: TextStyle }>,
    defaultStyle: TextStyle = {},
  ): void {
    textRuns.forEach((run) => {
      const r = this.document.createElement('a:r');
      p.appendChild(r);

      const rPr = this.document.createElement('a:rPr');
      r.appendChild(rPr);

      // Apply default styling first, then override with text run specific styling
      const mergedStyle = this.mergeStyles(defaultStyle, run.style);
      if (mergedStyle) {
        ModifyTextHelper.style(mergedStyle)(rPr);

        // Apply hyperlink if present
        if (mergedStyle.hyperlink) {
          this.applyHyperlink(rPr, mergedStyle.hyperlink);
        }
      }

      this.createTextElement(r, run.text || '');
    });
  }

  /**
   * Create a single text run with the given text and style
   */
  private createSingleTextRun(
    p: XmlElement,
    text: string | number,
    style?: TextStyle,
  ): void {
    const r = this.document.createElement('a:r');
    p.appendChild(r);

    const rPr = this.document.createElement('a:rPr');
    r.appendChild(rPr);

    // Apply text styling
    if (style) {
      ModifyTextHelper.style(style)(rPr);

      // Apply hyperlink if present
      if (style.hyperlink) {
        this.applyHyperlink(rPr, style.hyperlink);
      }
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

  /**
   * Apply hyperlink to a text run properties element
   */
  private applyHyperlink(
    rPr: XmlElement,
    hyperlinkInfo: { url: string; isInternal?: boolean; slideNumber?: number },
  ): void {
    if (!this.relationElement) {
      Logger.log(
        'MultiTextHelper: Cannot create hyperlink - no relation element provided',
        1,
      );
      return;
    }

    // Create relationship
    const relData = this.createRelationshipData(hyperlinkInfo);
    const newRelId = this.addRelationship(relData);

    // Create and append hyperlink element
    const hyperlinkElement = new HyperlinkElement(
      this.document,
      newRelId,
      hyperlinkInfo.isInternal || false,
    );

    rPr.appendChild(hyperlinkElement.createHlinkClick());
  }

  /**
   * Create relationship data for hyperlink
   */
  private createRelationshipData(hyperlinkInfo: {
    url: string;
    isInternal?: boolean;
    slideNumber?: number;
  }): {
    Type: string;
    Target: string;
    TargetMode?: string;
  } {
    if (hyperlinkInfo.isInternal) {
      return {
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
        Target: `../slides/${hyperlinkInfo.url}`,
      };
    }

    return {
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      Target: hyperlinkInfo.url,
      TargetMode: 'External',
    };
  }

  /**
   * Add relationship and return the relationship ID
   */
  private addRelationship(relData: {
    Type: string;
    Target: string;
    TargetMode?: string;
  }): string {
    const relNodes = this.relationElement.getElementsByTagName('Relationship');
    const maxId = XmlHelper.getMaxId(relNodes, 'Id', true);
    const newRelId = `rId${maxId}`;

    const newRel = this.relationElement.ownerDocument.createElement(
      'Relationship',
    );
    newRel.setAttribute('Id', newRelId);
    newRel.setAttribute('Type', relData.Type);
    newRel.setAttribute('Target', relData.Target);
    if (relData.TargetMode) {
      newRel.setAttribute('TargetMode', relData.TargetMode);
    }

    // Append to the root element of the relationships document
    const relRoot = relNodes.item(0)?.parentNode;
    if (relRoot) {
      relRoot.appendChild(newRel);
    }

    return newRelId;
  }
}
