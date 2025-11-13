import { HyperlinkInfo, TextStyle } from '../types/modify-types';
import { XmlElement } from '../types/xml-types';
import { MultiTextParagraph } from '../interfaces/imulti-text';
import ModifyTextHelper from './modify-text-helper';
import { XmlHelper } from './xml-helper';
import HyperlinkElement from './modify-hyperlink-element';
import { Logger, vd } from './general-helper';

type TextRun = { text: string; style?: TextStyle; hyperlink?: HyperlinkInfo };

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
    textRuns: TextRun[],
    defaultStyle: TextStyle = {},
  ): void {
    textRuns.forEach((run) => {
      const mergedStyle = this.mergeStyles(defaultStyle, run.style);

      // Check if this text run needs a hyperlink
      if (mergedStyle?.hyperlink && this.relationElement) {
        this.createHyperlinkTextRun(p, run.text || '', mergedStyle);
      } else {
        this.createRegularTextRun(p, run.text || '', mergedStyle);
      }
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
    const textString = String(text || '');

    // Check if this text run needs a hyperlink
    if (style?.hyperlink && this.relationElement) {
      this.createHyperlinkTextRun(p, textString, style);
    } else {
      this.createRegularTextRun(p, textString, style);
    }
  }

  /**
   * Create a regular text run without hyperlink
   */
  private createRegularTextRun(
    p: XmlElement,
    text: string,
    style?: TextStyle,
  ): void {
    const r = this.document.createElement('a:r');
    p.appendChild(r);

    const rPr = this.document.createElement('a:rPr');
    r.appendChild(rPr);

    // Apply text styling (excluding hyperlink)
    if (style) {
      const styleWithoutHyperlink = { ...style };
      delete styleWithoutHyperlink.hyperlink;
      ModifyTextHelper.style(styleWithoutHyperlink)(rPr);
    }

    this.createTextElement(r, text);
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
   * Create a hyperlinked text run using HyperlinkElement
   */
  private createHyperlinkTextRun(
    p: XmlElement,
    text: string,
    style: TextStyle,
  ): void {
    if (!style.hyperlink || !this.relationElement) {
      // Fallback to regular text run if no hyperlink info or relation element
      this.createRegularTextRun(p, text, style);
      return;
    }

    const hyperlinkInfo = style.hyperlink;
    const target = hyperlinkInfo.target;
    const isInternal = hyperlinkInfo.isInternal;

    if (!target) {
      this.createRegularTextRun(p, text, style);
      return;
    }

    // Create relationship data for the hyperlink
    const relData = this.createRelationshipData(target, isInternal || false);
    const newRelId = this.addRelationship(this.relationElement, relData);

    // Create HyperlinkElement to generate the hyperlinked text run
    const hyperlinkElement = new HyperlinkElement(
      this.document,
      newRelId,
      isInternal || false,
    );

    // Create the text run with hyperlink
    const r = hyperlinkElement.createTextRun(text);

    // Apply additional styling if needed (excluding hyperlink)
    if (style) {
      const rPr = r.getElementsByTagName('a:rPr')[0];
      if (rPr) {
        const styleWithoutHyperlink = { ...style };
        delete styleWithoutHyperlink.hyperlink;
        ModifyTextHelper.style(styleWithoutHyperlink)(rPr);
      }
    }

    p.appendChild(r);
  }

  /**
   * Create relationship data for hyperlink (reused from ModifyHyperlinkHelper)
   */
  private createRelationshipData(
    target: string | number,
    isInternal: boolean,
  ): { Id: string; Target: string; Type: string; TargetMode?: string } {
    if (isInternal) {
      return {
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
        Target: `../slides/${this.formatTarget(target)}`,
        Id: '', // Will be set later
      };
    }

    return {
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      Target: target.toString(),
      TargetMode: 'External',
      Id: '', // Will be set later
    };
  }

  private formatTarget(target: HyperlinkInfo['target']) {
    // For internal links, ensure the target is properly formatted
    let formattedTarget = target;
    if (typeof target === 'number') {
      formattedTarget = `slide${target}.xml`;
    } else if (typeof target === 'string' && !target.includes('.xml')) {
      // If it's a string but doesn't end with .xml, assume it's a slide number
      formattedTarget = `slide${target}.xml`;
    }
    return formattedTarget
  }

  /**
   * Add relationship to the relation element (reused logic from ModifyHyperlinkHelper)
   */
  private addRelationship(
    relation: XmlElement,
    relData: { Id: string; Target: string; Type: string; TargetMode?: string },
  ): string {
    const relNodes = relation.getElementsByTagName('Relationship');
    const maxId = XmlHelper.getMaxId(relNodes, 'Id', true);

    const newRelId = `rId${maxId}`;

    const newRel = relation.ownerDocument.createElement('Relationship');
    newRel.setAttribute('Id', newRelId);
    Object.entries(relData).forEach(([key, value]) => {
      if (value) newRel.setAttribute(key, value);
    });

    relNodes.item(0).parentNode.appendChild(newRel);

    return newRelId;
  }
}
