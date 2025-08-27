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
    if (paragraphProps.alignment) {
      const algn = this.document.createElement('a:algn');
      algn.setAttribute('val', paragraphProps.alignment);
      pPr.appendChild(algn);
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


const flattenAndParseHtml = (html: string): MultiTextParagraph[] => {
  const paragraphs: MultiTextParagraph[] = [];

  // Use DOMParser to parse the HTML
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  // Process top-level elements
  const topLevelElements = doc.body.children;

  // Track bullet list level
  let currentBulletLevel = 0;

  // Helper function to process text nodes and create textRuns
  const processTextNode = (
    node: Node,
    style: TextStyle = {},
  ): { text: string; style?: TextStyle } => {
    // If this is a text node, return its content
    if (node.nodeType === Node.TEXT_NODE) {
      return { text: node.textContent || '', style };
    }

    // If this is an element, handle specific styling
    if (node.nodeType === Node.ELEMENT_NODE) {
      const element = node as Element;
      const newStyle = { ...style };

      // Handle styling based on element type
      if (element.tagName === 'STRONG' || element.tagName === 'B') {
        newStyle.isBold = true;
      } else if (element.tagName === 'EM' || element.tagName === 'I') {
        newStyle.isItalics = true;
      } else if (element.tagName === 'SPAN') {
        // Process inline style attributes
        const styleAttr = element.getAttribute('style');
        if (styleAttr) {
          // Extract font size
          const fontSizeMatch = styleAttr.match(/font-size:\s*(\d+)px/i);
          if (fontSizeMatch && fontSizeMatch[1]) {
            newStyle.size = parseInt(fontSizeMatch[1]) * 100; // Convert px to points (100ths of point)
          }

          // Extract color (basic implementation)
          const colorMatch = styleAttr.match(/color:\s*([^;]+)/i);
          if (colorMatch && colorMatch[1]) {
            newStyle.color = {
              type: 'srgbClr',
              value: colorMatch[1].trim(),
            };
          }
        }
      }

      // For leaf nodes (no children), just return the text content with style
      if (element.childNodes.length === 0) {
        return { text: element.textContent || '', style: newStyle };
      }

      // For nodes with children, recursively process children and merge
      // This handles nested styling (e.g., <strong><em>text</em></strong>)
      const runs: { text: string; style?: TextStyle }[] = [];
      element.childNodes.forEach((child) => {
        const childRun = processTextNode(child, newStyle);
        if (childRun.text) {
          runs.push(childRun);
        }
      });

      // If we have a single run, return it directly
      if (runs.length === 1) {
        return runs[0];
      }

      // If we have multiple runs, merge adjacent runs with the same style
      // For simplicity in this implementation, we'll just return a concatenated text
      return {
        text: runs.map((run) => run.text).join(''),
        style: newStyle,
      };
    }

    return { text: '' };
  };

  // Process a node and its children to create paragraphs
  const processNode = (node: Element, level = 0) => {
    const tagName = node.tagName.toLowerCase();

    // Handle different node types
    if (tagName === 'p') {
      // Create a paragraph
      const textRuns: { text: string; style?: TextStyle }[] = [];

      // Process all child nodes to create text runs
      node.childNodes.forEach((child) => {
        const run = processTextNode(child);
        if (run.text) {
          textRuns.push(run);
        }
      });

      // If no text runs were created, add an empty one
      if (textRuns.length === 0) {
        textRuns.push({ text: '' });
      }

      // Create the paragraph
      paragraphs.push({
        paragraph: {
          level: 0,
          bullet: false,
          alignment: 'left',
        },
        textRuns,
      });
    } else if (tagName === 'ul' || tagName === 'ol') {
      // Increase bullet level for nested lists
      currentBulletLevel++;

      // Process all list items
      Array.from(node.children).forEach((child) => {
        if (child.tagName.toLowerCase() === 'li') {
          processNode(child, currentBulletLevel);
        } else {
          // Handle nested lists
          processNode(child, currentBulletLevel);
        }
      });

      // Decrease bullet level after processing list
      currentBulletLevel--;
    } else if (tagName === 'li') {
      // Create a paragraph for the list item
      const textRuns: { text: string; style?: TextStyle }[] = [];

      // Process all child nodes to create text runs
      let textContent = '';
      let hasNestedList = false;

      Array.from(node.childNodes).forEach((child) => {
        // Skip nested lists, they'll be processed separately
        if (
          child.nodeType === Node.ELEMENT_NODE &&
          (child.nodeName.toLowerCase() === 'ul' ||
            child.nodeName.toLowerCase() === 'ol')
        ) {
          hasNestedList = true;
          return;
        }

        const run = processTextNode(child);
        if (run.text) {
          textRuns.push(run);
        }
      });

      // If no text runs were created, add an empty one
      if (textRuns.length === 0) {
        textRuns.push({ text: '' });
      }

      // Create the paragraph for the list item
      paragraphs.push({
        paragraph: {
          level,
          bullet: true,
          alignment: 'left',
        },
        textRuns,
      });

      // Process nested lists if any
      Array.from(node.children).forEach((child) => {
        if (
          child.tagName.toLowerCase() === 'ul' ||
          child.tagName.toLowerCase() === 'ol'
        ) {
          processNode(child, level + 1);
        }
      });
    } else {
      // For other elements, process their children
      Array.from(node.children).forEach((child) => {
        processNode(child, level);
      });
    }
  };

  // Process all top-level elements
  Array.from(topLevelElements).forEach((element) => {
    processNode(element);
  });

  return paragraphs;
};
