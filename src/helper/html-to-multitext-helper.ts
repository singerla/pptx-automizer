import { TextStyle } from '../types/modify-types';
import {
  MultiTextAutoNumberType,
  MultiTextParagraph,
  MultiTextRun,
} from '../interfaces/imulti-text';
import { DOMParser, Node } from '@xmldom/xmldom';
import type { Element } from '@xmldom/xmldom';
import { log } from './logger';
import {
  normalizeCssColor,
  parseCssFontFamily,
  parseCssFontSize,
  parseCssFontStyle,
  parseCssFontWeight,
  parseCssTextAlign,
  parseCssTextDecoration,
  parseInlineCss,
} from './css-style-parser';

type Alignment = MultiTextParagraph['paragraph']['alignment'];

/**
 * Everything a paragraph inherits from its *block* ancestors. Block elements
 * never nest in the output: an inner block flushes the current paragraph and
 * opens a new one carrying these properties.
 */
type BlockContext = {
  /** List nesting stack, one entry per open <ul>/<ol> ('ul' | 'ol') */
  listTypes: ('ul' | 'ol')[];
  alignment?: Alignment;
  /** Style contributed by the block itself (e.g. h1 size, pre monospace) */
  blockStyle: TextStyle;
  /** The block (or an ancestor) carries a default vertical margin (<p>, h1…) */
  margin?: boolean;
  /**
   * Identity of the top-level list this block sits in. Items of the same list
   * sit tight; where two *different* lists touch, each list's own outer
   * margin applies.
   */
  listRoot?: number;
};

/** The paragraph currently being filled. */
type OpenParagraph = {
  block: BlockContext;
  runs: MultiTextRun[];
};

/**
 * Tags whose empty form is a deliberate blank line rather than a wrapper:
 * `<p></p>` and `<p>&nbsp;</p>` are how editors express "empty line", while an
 * empty `<div>` is just structure.
 */
const PRESERVE_EMPTY_TAGS = ['p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];

/** Inline tags that only ever contribute character properties. */
const INLINE_TAGS = [
  'a',
  'abbr',
  'b',
  'big',
  'cite',
  'code',
  'del',
  'em',
  'font',
  'i',
  'ins',
  'kbd',
  'mark',
  'q',
  's',
  'samp',
  'small',
  'span',
  'strike',
  'strong',
  'sub',
  'sup',
  'tt',
  'u',
  'var',
];

/** Block tags that open a paragraph of their own. */
const BLOCK_TAGS = [
  'address',
  'article',
  'aside',
  'blockquote',
  'div',
  'dd',
  'dt',
  'figcaption',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'header',
  'footer',
  'main',
  'p',
  'pre',
  'section',
];

/** Headings: paragraph level styling, relative to a 12pt body default. */
const HEADING_SIZES: Record<string, number> = {
  h1: 2400,
  h2: 2000,
  h3: 1800,
  h4: 1600,
  h5: 1400,
  h6: 1200,
};

const MONOSPACE_FONT = 'Consolas';

/**
 * Blocks that carry a vertical margin in the browser's default stylesheet
 * (`margin: 1em 0`); `<ul>`/`<ol>` carry one too, handled as list edges.
 */
const MARGIN_TAGS = [
  'p',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'blockquote',
  'pre',
];

/**
 * Gap between block paragraphs as a percentage of the line height - the PPTX
 * projection of the browser's `margin: 1em 0`, with adjacent margins already
 * collapsed into one gap.
 */
const BLOCK_MARGIN_PERCENT = 100;

/**
 * PowerPoint varies the numbering scheme per outline level; mirror the
 * familiar 1. / a. / i. cascade and cycle beyond level 3.
 */
const AUTO_NUMBER_TYPES: MultiTextAutoNumberType[] = [
  'arabicPeriod',
  'alphaLcPeriod',
  'romanLcPeriod',
];

export class HtmlToMultiTextHelper {
  private paragraphs: MultiTextParagraph[] = [];
  /** Parallel to `paragraphs`: margin/list identity of the emitting block. */
  private blockMeta: { margin: boolean; listRoot?: number }[] = [];
  private listCounter = 0;
  private open?: OpenParagraph;

  /**
   * Converts an HTML string to a MultiTextParagraph array.
   *
   * The input contract is XHTML-ish markup as produced by WYSIWYG editors
   * (CKEditor/TinyMCE): tags closed, attributes quoted. `@xmldom/xmldom` is an
   * XML parser - real-world tag soup (unclosed `<br>`, bare `&nbsp;`) does not
   * parse. Common editor quirks that *are* handled: sibling-nested lists
   * (`<ul><li/><ul>…</ul></ul>`), block-inside-block, and inline styles on any
   * element.
   *
   * @param html HTML string to convert
   * @returns Array of MultiTextParagraph objects
   */
  public run(html: string): MultiTextParagraph[] {
    this.paragraphs = [];
    this.blockMeta = [];
    this.listCounter = 0;
    this.open = undefined;

    const doc = new DOMParser().parseFromString(html, 'text/html');
    const body = doc.getElementsByTagName('body')[0];

    if (!body) {
      log.error('You need to provide a <body> tag for HtmlToMultiText');
      return this.paragraphs;
    }

    const rootBlock: BlockContext = { listTypes: [], blockStyle: {} };

    this.walkChildren(body as unknown as Element, rootBlock, {});
    this.flush();
    this.applyBlockMargins();

    return this.paragraphs;
  }

  /**
   * Single traversal pass. Two accumulators travel down the tree:
   * `block` (list depth, alignment, block styling) and `inline` (character
   * properties). Neither is ever mutated in place - each level gets its own
   * copy, which is what makes both list-nesting shapes converge.
   */
  private walk(node: Node, block: BlockContext, inline: TextStyle): void {
    if (node.nodeType === Node.TEXT_NODE) {
      this.appendText(node.textContent || '', block, inline);
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const element = node as Element;
    const tagName = element.tagName.toLowerCase();

    switch (true) {
      case tagName === 'br':
        this.appendBreak(block, inline);
        return;

      case tagName === 'ul' || tagName === 'ol':
        // A list interrupts its parent paragraph: <li>text<ul>… must not merge
        this.flush();
        this.walkChildren(
          element,
          this.pushList(element, block, tagName),
          inline,
        );
        return;

      case tagName === 'li':
      case BLOCK_TAGS.includes(tagName):
        this.walkBlock(element, tagName, block, inline);
        return;

      case INLINE_TAGS.includes(tagName):
        this.walkChildren(element, block, this.inlineStyleFor(element, tagName, inline));
        return;

      default:
        // Unknown/structural elements (table, span-like customs, …): keep their
        // text instead of dropping it
        this.walkChildren(element, block, inline);
    }
  }

  /**
   * A block element (incl. `<li>`) opens exactly one paragraph. If its children
   * turn out to be blocks themselves, *they* become the paragraphs and this one
   * stays empty - the innermost block wins, ancestors only contribute
   * properties. Only a block that emitted nothing at all may end up as a
   * deliberate blank line.
   */
  private walkBlock(
    element: Element,
    tagName: string,
    block: BlockContext,
    inline: TextStyle,
  ): void {
    // A block starts a paragraph of its own - whatever text was collected
    // before it (loose text, a preceding inline run) ends here
    this.flush();

    const blockContext = this.blockContextFor(element, tagName, block);
    const paragraphsBefore = this.paragraphs.length;

    this.openParagraph(blockContext);
    this.walkChildren(element, blockContext, inline);

    const emittedNothing = this.paragraphs.length === paragraphsBefore;
    this.flush(PRESERVE_EMPTY_TAGS.includes(tagName) && emittedNothing);
  }

  private walkChildren(
    element: Element,
    block: BlockContext,
    inline: TextStyle,
  ): void {
    Array.from(element.childNodes).forEach((child) =>
      this.walk(child as Node, block, inline),
    );
  }

  /** Push one list level, taking the list type from the tag. */
  private pushList(
    element: Element,
    block: BlockContext,
    tagName: 'ul' | 'ol',
  ): BlockContext {
    const withCss = this.applyElementCss(element, block);

    return {
      ...withCss,
      listTypes: [...withCss.listTypes, tagName],
      // A new top-level list gets its own identity; nested lists inherit it
      listRoot:
        withCss.listTypes.length === 0 ? ++this.listCounter : withCss.listRoot,
    };
  }

  /** Block context of a non-list block element, incl. heading/pre styling. */
  private blockContextFor(
    element: Element,
    tagName: string,
    block: BlockContext,
  ): BlockContext {
    const context = this.applyElementCss(element, block);
    const blockStyle: TextStyle = { ...context.blockStyle };

    if (HEADING_SIZES[tagName]) {
      blockStyle.size = HEADING_SIZES[tagName];
      blockStyle.isBold = true;
    }

    if (tagName === 'pre') {
      blockStyle.fontFamily = MONOSPACE_FONT;
    }

    return {
      ...context,
      blockStyle,
      margin: context.margin || MARGIN_TAGS.includes(tagName),
    };
  }

  /** Merge an element's inline CSS into the block context (alignment, style). */
  private applyElementCss(element: Element, block: BlockContext): BlockContext {
    const css = parseInlineCss(element.getAttribute('style') || '');
    if (Object.keys(css).length === 0) {
      return block;
    }

    const alignment = css['text-align']
      ? parseCssTextAlign(css['text-align'])
      : undefined;

    return {
      ...block,
      alignment: alignment ?? block.alignment,
      blockStyle: this.cssToTextStyle(css, block.blockStyle),
    };
  }

  /** Character properties contributed by one inline element. */
  private inlineStyleFor(
    element: Element,
    tagName: string,
    inline: TextStyle,
  ): TextStyle {
    const style: TextStyle = { ...inline };

    switch (tagName) {
      case 'strong':
      case 'b':
        style.isBold = true;
        break;
      case 'em':
      case 'i':
      case 'cite':
      case 'var':
        style.isItalics = true;
        break;
      case 'u':
      case 'ins':
        style.isUnderlined = true;
        break;
      case 's':
      case 'strike':
      case 'del':
        style.isStrike = true;
        break;
      case 'sub':
        style.isSubscript = true;
        break;
      case 'sup':
        style.isSuperscript = true;
        break;
      case 'code':
      case 'kbd':
      case 'samp':
      case 'tt':
        style.fontFamily = MONOSPACE_FONT;
        break;
      case 'mark':
        style.highlight = { type: 'srgbClr', value: 'FFFF00' };
        break;
      case 'a':
        this.applyHyperlink(element, style);
        break;
      case 'font':
        this.applyFontAttributes(element, style);
        break;
    }

    return this.cssToTextStyle(
      parseInlineCss(element.getAttribute('style') || ''),
      style,
    );
  }

  /** Map the supported CSS subset onto TextStyle. */
  private cssToTextStyle(
    css: Record<string, string>,
    base: TextStyle,
  ): TextStyle {
    if (Object.keys(css).length === 0) {
      return base;
    }

    const style: TextStyle = { ...base };

    const size = css['font-size']
      ? parseCssFontSize(css['font-size'])
      : undefined;
    if (size !== undefined) {
      style.size = size;
    }

    const color = css['color'] ? normalizeCssColor(css['color']) : undefined;
    if (color) {
      style.color = { type: 'srgbClr', value: color };
    }

    const highlight = css['background-color']
      ? normalizeCssColor(css['background-color'])
      : undefined;
    if (highlight) {
      style.highlight = { type: 'srgbClr', value: highlight };
    }

    const isBold = css['font-weight']
      ? parseCssFontWeight(css['font-weight'])
      : undefined;
    if (isBold !== undefined) {
      style.isBold = isBold;
    }

    const isItalics = css['font-style']
      ? parseCssFontStyle(css['font-style'])
      : undefined;
    if (isItalics !== undefined) {
      style.isItalics = isItalics;
    }

    const decoration =
      css['text-decoration'] ?? css['text-decoration-line'] ?? undefined;
    if (decoration !== undefined) {
      const { isUnderlined, isStrike } = parseCssTextDecoration(decoration);
      if (isUnderlined !== undefined) {
        style.isUnderlined = isUnderlined;
      }
      if (isStrike !== undefined) {
        style.isStrike = isStrike;
      }
    }

    const fontFamily = css['font-family']
      ? parseCssFontFamily(css['font-family'])
      : undefined;
    if (fontFamily) {
      style.fontFamily = fontFamily;
    }

    return style;
  }

  /**
   * `<a href>`: a purely numeric href is treated as an internal slide number,
   * anything else as an external target.
   */
  private applyHyperlink(element: Element, style: TextStyle): void {
    const href = element.getAttribute('href');
    if (!href) return;

    if (/^\d+$/.test(href.trim())) {
      style.hyperlink = {
        target: parseInt(href, 10),
        isInternal: true,
      };
      return;
    }

    style.hyperlink = {
      target: href,
      isInternal: false,
    };
  }

  /** Legacy `<font color size face>` - still emitted by some editors. */
  private applyFontAttributes(element: Element, style: TextStyle): void {
    const color = element.getAttribute('color');
    if (color) {
      const normalized = normalizeCssColor(color);
      if (normalized) {
        style.color = { type: 'srgbClr', value: normalized };
      }
    }

    const face = element.getAttribute('face');
    if (face) {
      const fontFamily = parseCssFontFamily(face);
      if (fontFamily) {
        style.fontFamily = fontFamily;
      }
    }
  }

  /**
   * Append text to the open paragraph, applying CSS whitespace collapsing:
   * runs of whitespace become a single space, and a space is dropped at the
   * start of a paragraph. `&nbsp;` (U+00A0) is not whitespace and survives.
   */
  private appendText(
    text: string,
    block: BlockContext,
    inline: TextStyle,
  ): void {
    const collapsed = text.replace(/[ \t\r\n\f]+/g, ' ');

    if (collapsed === '') {
      return;
    }

    // Whitespace between block elements is layout, not content
    if (collapsed === ' ' && !this.open) {
      return;
    }

    this.openParagraph(block);

    const isParagraphStart = this.open.runs.length === 0;
    const text_ = isParagraphStart ? collapsed.replace(/^ /, '') : collapsed;

    if (text_ === '') {
      return;
    }

    this.open.runs.push({ text: text_, style: this.runStyle(block, inline) });
  }

  private appendBreak(block: BlockContext, inline: TextStyle): void {
    this.openParagraph(block);
    this.open.runs.push({ break: true, style: this.runStyle(block, inline) });
  }

  /** Block styling is the base, inline styling wins. */
  private runStyle(block: BlockContext, inline: TextStyle): TextStyle {
    return { ...block.blockStyle, ...inline };
  }

  private openParagraph(block: BlockContext): void {
    if (this.open) {
      return;
    }
    this.open = { block, runs: [] };
  }

  /**
   * Close the open paragraph and project its block context onto the flat PPTX
   * paragraph properties.
   *
   * @param allowEmpty - Emit the paragraph even without runs (blank line)
   */
  private flush(allowEmpty = false): void {
    if (!this.open) {
      return;
    }

    const { block, runs } = this.open;
    this.open = undefined;

    this.trimTrailingWhitespace(runs);

    // A <br> right before the block's end has no visible effect in HTML (the
    // empty last line gets no line box) - drop exactly one, so a deliberate
    // <br><br> still keeps its one empty line
    if (runs[runs.length - 1]?.break) {
      runs.pop();
      this.trimTrailingWhitespace(runs);
    }

    if (runs.length === 0) {
      if (!allowEmpty) {
        return;
      }
      runs.push({ text: '' });
    }

    const listDepth = block.listTypes.length;
    const listType = block.listTypes[listDepth - 1];
    // PPTX `lvl` is 0-based: the outermost list level is 0, like a plain
    // paragraph - the bullet, not the level, is what sets a list item apart
    const level = listDepth > 0 ? Math.min(listDepth - 1, 8) : 0;

    this.paragraphs.push({
      paragraph: {
        level,
        bullet: listDepth > 0,
        ...(listType === 'ol'
          ? {
              bulletType: 'number' as const,
              autoNumberType:
                AUTO_NUMBER_TYPES[level % AUTO_NUMBER_TYPES.length],
            }
          : {}),
        // Only set when the HTML says so: an unaligned paragraph inherits from
        // the shape's layout, the same way HTML inherits from its container
        ...(block.alignment ? { alignment: block.alignment } : {}),
      },
      textRuns: runs,
    });
    this.blockMeta.push({ margin: !!block.margin, listRoot: block.listRoot });
  }

  /** Trailing collapsed whitespace is layout, not content. */
  private trimTrailingWhitespace(runs: MultiTextRun[]): void {
    for (let i = runs.length - 1; i >= 0; i--) {
      const run = runs[i];
      if (run.text === undefined) {
        return;
      }
      run.text = run.text.replace(/ +$/, '');
      if (run.text !== '') {
        return;
      }
      runs.pop();
    }
  }

  /**
   * Project the browser's default vertical margins onto `spaceBefore`.
   *
   * Items of the same list sit tight (nested lists carry no margin in the
   * default stylesheet); every other boundary touching a margin-bearing block
   * or a list edge - including the edge between two adjacent lists - gets one
   * collapsed gap, carried by the later paragraph. The first paragraph gets
   * none: there is no margin against the text frame.
   */
  private applyBlockMargins(): void {
    this.paragraphs.forEach((paragraph, index) => {
      if (index === 0) {
        return;
      }

      const previous = this.blockMeta[index - 1];
      const current = this.blockMeta[index];
      const sameList =
        current.listRoot !== undefined &&
        current.listRoot === previous.listRoot;
      const touchesMargin =
        previous.margin ||
        current.margin ||
        previous.listRoot !== undefined ||
        current.listRoot !== undefined;

      if (!sameList && touchesMargin) {
        paragraph.paragraph.spaceBefore = { percent: BLOCK_MARGIN_PERCENT };
      }
    });
  }
}
