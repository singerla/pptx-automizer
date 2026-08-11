import { HyperlinkInfo, TextStyle } from '../types/modify-types';
import { XmlElement } from '../types/xml-types';
import { MultiTextParagraph, MultiTextRun } from '../interfaces/imulti-text';
import ModifyTextHelper from './modify-text-helper';
import { XmlHelper } from './xml-helper';
import HyperlinkElement from './modify-hyperlink-element';

type TextRun = MultiTextRun & { hyperlink?: HyperlinkInfo };

/**
 * Children of <a:pPr>, in OOXML schema order (CT_TextParagraphProperties).
 * Spacing comes before the bullet properties: appending bullets first is what
 * used to make PowerPoint offer to repair the file.
 */
const PPR_CHILD_ORDER = [
  'a:lnSpc',
  'a:spcBef',
  'a:spcAft',
  'a:buClrTx',
  'a:buClr',
  'a:buSzTx',
  'a:buSzPct',
  'a:buSzPts',
  'a:buFontTx',
  'a:buFont',
  'a:buNone',
  'a:buAutoNum',
  'a:buChar',
  'a:tabLst',
  'a:defRPr',
  'a:extLst',
];

/** Default bullet indent step, approx. 0.25 inch in EMU. */
const BULLET_INDENT = 228600;

/** Typeface carrying the default '•' glyph. */
const BULLET_FONT = 'Arial';

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
   * Extract the fallback style from the existing text, so generated runs keep
   * the template's look where the caller says nothing.
   *
   * Restricted to size and color on purpose: bold/italic of the template's
   * first run are a property of *that* text, not of the shape, and inheriting
   * them made every generated run bold as soon as the placeholder was
   * (e.g. a title). Callers who want bold pass it in their TextStyle.
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

      }
    }

    return defaultStyle;
  }

  /**
   * Find or create the txBody element and ensure it has required child elements
   *
   * Notes:
   * - Text boxes / shapes use `p:txBody`
   * - Table cells (`a:tc`) use `a:txBody`
   */
  private getOrCreateTxBody(): XmlElement {
    // 1) Find existing txBody for both shapes and table cells
    let txBody =
      this.element.getElementsByTagName('p:txBody')[0] ||
      this.element.getElementsByTagName('a:txBody')[0];

    // 2) If missing, create correct txBody depending on host element
    if (!txBody) {
      const isTableCell = this.element.tagName === 'a:tc';

      // Table cell text lives in DrawingML (`a:*`), shape text in PresentationML (`p:*`)
      txBody = this.document.createElement(isTableCell ? 'a:txBody' : 'p:txBody');
      this.element.appendChild(txBody);
    }

    // 3) Ensure required children exist (always DrawingML children)
    let bodyPr = txBody.getElementsByTagName('a:bodyPr')[0];
    if (!bodyPr) {
      bodyPr = this.document.createElement('a:bodyPr');
      txBody.insertBefore(bodyPr, txBody.firstChild);
    }

    let lstStyle = txBody.getElementsByTagName('a:lstStyle')[0];
    if (!lstStyle) {
      lstStyle = this.document.createElement('a:lstStyle');
      const insertAfter = bodyPr.nextSibling;
      if (insertAfter) {
        txBody.insertBefore(lstStyle, insertAfter);
      } else {
        txBody.appendChild(lstStyle);
      }
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
  private applyParagraphProperties(
    pPr: XmlElement,
    paragraphProps: MultiTextParagraph['paragraph'],
  ): void {
    const level = paragraphProps.level ?? 0;

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
    } else if (paragraphProps.bullet) {
      // Hanging indent, so wrapped lines align with the text, not the bullet
      pPr.setAttribute('indent', String(-BULLET_INDENT));
    }

    // Set left margin
    if (paragraphProps.marginLeft !== undefined) {
      pPr.setAttribute('marL', String(paragraphProps.marginLeft));
    } else if (paragraphProps.bullet) {
      // Indent per level: a constant margin renders every nested bullet at the
      // same x-position unless the target layout happens to define a lstStyle
      pPr.setAttribute('marL', String((level + 1) * BULLET_INDENT));
    }

    // Apply spacing properties
    this.applySpacingProperties(pPr, paragraphProps);

    // The steps above append in call order; the schema wants a fixed sequence
    XmlHelper.sortChildrenBySchema(pPr, PPR_CHILD_ORDER);
  }

  /**
   * Apply bullet configuration to paragraph properties.
   *
   * PPTX has no list object - a list item is a paragraph with bullet
   * properties, either a literal glyph (`a:buChar`) or automatic numbering
   * (`a:buAutoNum`).
   */
  private applyBulletConfiguration(
    pPr: XmlElement,
    paragraphProps: MultiTextParagraph['paragraph'],
  ): void {
    if (paragraphProps.bullet) {
      if (paragraphProps.bulletType === 'number') {
        const buAutoNum = this.document.createElement('a:buAutoNum');
        buAutoNum.setAttribute(
          'type',
          paragraphProps.autoNumberType || 'arabicPeriod',
        );
        pPr.appendChild(buAutoNum);
        return;
      }

      // A glyph bullet needs the typeface that actually carries the glyph
      const buFont = this.document.createElement('a:buFont');
      buFont.setAttribute('typeface', BULLET_FONT);
      pPr.appendChild(buFont);

      const buChar = this.document.createElement('a:buChar');
      buChar.setAttribute('char', paragraphProps.bulletChar || '•');
      pPr.appendChild(buChar);
      return;
    }

    if (paragraphProps.bullet === false || paragraphProps.level === 0) {
      // Suppress a bullet the layout's lstStyle would otherwise inherit
      const buNone = this.document.createElement('a:buNone');
      pPr.appendChild(buNone);
    }
  }

  /**
   * Apply spacing properties to paragraph properties
   */
  private applySpacingProperties(
    pPr: XmlElement,
    paragraphProps: MultiTextParagraph['paragraph'],
  ): void {
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

      // An explicit break is an element between runs, not a run with text
      if (run.break) {
        this.appendLineBreak(p, mergedStyle);
        return;
      }

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
    this.forEachLine(p, text, style, (line) => {
      const r = this.document.createElement('a:r');
      p.appendChild(r);

      const rPr = this.document.createElement('a:rPr');
      r.appendChild(rPr);

      // Apply text styling (excluding hyperlink)
      this.applyRunStyle(rPr, style);

      this.createTextElement(r, line);
    });
  }

  /**
   * Split text at soft line breaks and emit an `<a:br/>` between the segments.
   *
   * PPTX has no in-run line break: a soft break (Shift+Enter) is an `<a:br/>`
   * element sitting *between* two `<a:r>` siblings. PowerPoint itself hands out
   * U+000B (vertical tab) for such a break when text is read back from a shape,
   * so `\v` is treated like `\n` here. A literal `\v` in `<a:t>` is not even
   * valid XML 1.0 and used to produce a file PowerPoint offers to repair.
   *
   * Empty segments (consecutive breaks) do not get a run of their own — only
   * text that was empty to begin with keeps its single empty run.
   */
  private forEachLine(
    p: XmlElement,
    text: string,
    style: TextStyle | undefined,
    createRun: (line: string) => void,
  ): void {
    // eslint-disable-next-line no-control-regex -- U+000B is one of the breaks we split on
    const lines = String(text ?? '').split(/\r\n|[\r\n\u000B]/);

    lines.forEach((line, index) => {
      if (index > 0) {
        this.appendLineBreak(p, style);
      }
      // Consecutive breaks: the break alone is the empty line
      if (line === '' && lines.length > 1) {
        return;
      }
      createRun(line);
    });
  }

  /**
   * Append an `<a:br/>` carrying the surrounding run style, so the empty line
   * keeps the font size (and therefore the line height) of its neighbours.
   */
  private appendLineBreak(p: XmlElement, style?: TextStyle): void {
    const br = this.document.createElement('a:br');
    const rPr = this.document.createElement('a:rPr');
    br.appendChild(rPr);
    this.applyRunStyle(rPr, style);
    p.appendChild(br);
  }

  /**
   * Apply a TextStyle to an `<a:rPr>`, ignoring the hyperlink part
   * (hyperlinks are an element inside rPr, added by HyperlinkElement)
   */
  private applyRunStyle(rPr: XmlElement, style?: TextStyle): void {
    if (!style) return;
    const styleWithoutHyperlink = { ...style };
    delete styleWithoutHyperlink.hyperlink;
    ModifyTextHelper.style(styleWithoutHyperlink)(rPr);
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
    const sanitized = XmlHelper.sanitizeText(text);
    const textNode = this.document.createTextNode(sanitized);
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

    // Create the text run(s) with hyperlink; a soft line break splits the text
    // into several runs, all pointing at the same relationship
    this.forEachLine(p, text, style, (line) => {
      const r = hyperlinkElement.createTextRun(line);

      // Apply additional styling if needed (excluding hyperlink)
      const rPr = r.getElementsByTagName('a:rPr')[0];
      if (rPr) {
        this.applyRunStyle(rPr as XmlElement, style);
      }

      p.appendChild(r);
    });
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
