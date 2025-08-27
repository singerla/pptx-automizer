import { TextStyle } from '../types/modify-types';
import { MultiTextParagraph } from '../interfaces/imulti-text';
import { DOMParser, Node } from '@xmldom/xmldom';

type TextRun = { text: string; style?: TextStyle };

export class HtmlToMultiTextHelper {
  /**
   * Converts HTML string to MultiTextParagraph array
   * @param html HTML string to convert
   * @returns Array of MultiTextParagraph objects
   */
  public run(html: string): MultiTextParagraph[] {
    const paragraphs: MultiTextParagraph[] = [];
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const currentBulletLevel = 0;

    // Get the body element using getElementsByTagName
    const bodyElement = doc.getElementsByTagName('body')[0];

    // Process all top-level elements
    if (bodyElement) {
      Array.from(bodyElement.childNodes).forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          this.processNode(
            node as unknown as Element,
            currentBulletLevel,
            paragraphs,
          );
        }
      });
    }

    return paragraphs;
  }

  /**
   * Processes an HTML node and converts it to MultiTextParagraph objects
   */
  private processNode(
    node: ChildNode,
    level = 0,
    paragraphs: MultiTextParagraph[],
    bulletLevel: { value: number } = { value: 0 },
  ): void {
    const tagName = node.nodeName.toLowerCase();

    switch (tagName) {
      case 'p':
        this.processParagraph(node, paragraphs);
        break;

      case 'ul':
      case 'ol':
        this.processList(node, paragraphs, bulletLevel);
        break;

      case 'li':
        this.processListItem(node, level, paragraphs);
        break;

      default:
        // For other elements, process their children
        Array.from(node.childNodes).forEach((child) => {
          this.processNode(child, level, paragraphs, bulletLevel);
        });
    }
  }

  /**
   * Processes a paragraph element
   */
  private processParagraph(
    node: ChildNode,
    paragraphs: MultiTextParagraph[],
  ): void {
    const textRuns = this.createTextRuns(node);

    // If no text runs were created, add an empty one
    if (textRuns.length === 0) {
      textRuns.push({ text: '' });
    }

    // Create the paragraph
    paragraphs.push({
      paragraph: {
        level: 0,
        bullet: false,
        alignment: 'l',
      },
      textRuns,
    });
  }

  /**
   * Processes a list (ul/ol) element
   */
  private processList(
    node: ChildNode,
    paragraphs: MultiTextParagraph[],
    bulletLevel: { value: number },
  ): void {
    // Increase bullet level for nested lists
    bulletLevel.value++;

    // Process all list items
    Array.from(node.childNodes).forEach((child) => {
      this.processNode(child, bulletLevel.value, paragraphs, bulletLevel);
    });

    // Decrease bullet level after processing list
    bulletLevel.value--;
  }

  /**
   * Processes a list item element
   */
  private processListItem(
    node: ChildNode,
    level: number,
    paragraphs: MultiTextParagraph[],
  ): void {
    const textRuns: TextRun[] = [];

    // Process all child nodes to create text runs (except nested lists)
    Array.from(node.childNodes).forEach((child) => {
      // Skip nested lists, they'll be processed separately
      if (
        child.nodeType === Node.ELEMENT_NODE &&
        (child.nodeName.toLowerCase() === 'ul' ||
          child.nodeName.toLowerCase() === 'ol')
      ) {
        return;
      }

      const run = this.processTextNode(child);
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
        alignment: 'l',
      },
      textRuns,
    });

    // Process nested lists if any
    Array.from(node.childNodes).forEach((child) => {
      if (
        child.nodeName.toLowerCase() === 'ul' ||
        child.nodeName.toLowerCase() === 'ol'
      ) {
        const bulletLevel = { value: level };
        this.processNode(child, level + 1, paragraphs, bulletLevel);
      }
    });
  }

  /**
   * Creates text runs from an element's child nodes
   */
  private createTextRuns(node: ChildNode): TextRun[] {
    const textRuns: TextRun[] = [];

    // Process all child nodes to create text runs
    // Using Array.from to convert NodeList to array that has forEach
    Array.from(node.childNodes).forEach((child) => {
      const run = this.processTextNode(child);
      if (run.text) {
        textRuns.push(run);
      }
    });

    return textRuns;
  }

  /**
   * Processes a text node and creates a TextRun
   */
  private processTextNode(node: ChildNode, style: TextStyle = {}): TextRun {
    // If this is a text node, return its content
    if (node.nodeType === Node.TEXT_NODE) {
      return { text: node.textContent || '', style };
    }

    // If this is an element, handle specific styling
    if (node.nodeType === Node.ELEMENT_NODE) {
      const element = node as Element;
      const newStyle = this.applyElementStyles(element, { ...style });

      // For leaf nodes (no children), just return the text content with style
      if (element.childNodes.length === 0) {
        return { text: element.textContent || '', style: newStyle };
      }

      // For nodes with children, recursively process children
      return this.processElementWithChildren(element, newStyle);
    }

    return { text: '' };
  }

  /**
   * Applies styles based on the element type and attributes
   */
  private applyElementStyles(element: Element, style: TextStyle): TextStyle {
    const newStyle = { ...style };

    const tagName = element.tagName.toLowerCase();

    // Handle styling based on element type
    if (tagName === 'strong' || tagName === 'b') {
      newStyle.isBold = true;
    } else if (tagName === 'em' || tagName === 'i') {
      newStyle.isItalics = true;
    } else if (tagName === 'span') {
      this.processSpanStyles(element, newStyle);
    }

    return newStyle;
  }

  /**
   * Processes span element styles
   */
  private processSpanStyles(element: Element, style: TextStyle): void {
    const styleAttr = element.getAttribute('style');
    if (!styleAttr) return;

    // Extract font size
    const fontSizeMatch = styleAttr.match(/font-size:\s*(\d+)px/i);
    if (fontSizeMatch && fontSizeMatch[1]) {
      style.size = parseInt(fontSizeMatch[1]) * 100; // Convert px to points (100ths of point)
    }

    // Extract color
    const colorMatch = styleAttr.match(/color:\s*([^;]+)/i);
    if (colorMatch && colorMatch[1]) {
      style.color = {
        type: 'srgbClr',
        value: colorMatch[1].trim(),
      };
    }
  }

  /**
   * Processes an element with child nodes
   */
  private processElementWithChildren(
    element: Element,
    style: TextStyle,
  ): TextRun {
    const runs: TextRun[] = [];

    Array.from(element.childNodes).forEach((child) => {
      const childRun = this.processTextNode(child, style);
      if (childRun.text) {
        runs.push(childRun);
      }
    });

    // If we have a single run, return it directly
    if (runs.length === 1) {
      return runs[0];
    }

    // If we have multiple runs, concatenate the text
    return {
      text: runs.map((run) => run.text).join(''),
      style,
    };
  }
}
