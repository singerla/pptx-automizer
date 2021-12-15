import { GeneralHelper, vd } from './general-helper';
import escape from 'regexp.escape';
import { XmlHelper } from './xml-helper';
import {
  ReplaceText,
  ReplaceTextOptions,
  TextStyle,
} from '../types/modify-types';
import ModifyTextHelper from './modify-text-helper';

type Expressions = {
  openingTag: string;
  closingTag: string;
};
type CharacterSplit = {
  from: number;
  to: number;
  text: string;
};

export default class TextReplaceHelper {
  expressions: Expressions;
  element: XMLDocument;
  newNodes: Element[];
  options: ReplaceTextOptions;

  constructor(options: ReplaceTextOptions, element: XMLDocument) {
    const defaultOptions = {
      openingTag: '{{',
      closingTag: '}}',
    };

    this.options = !options
      ? defaultOptions
      : { ...defaultOptions, ...options };

    this.element = element;
    this.expressions = {
      openingTag: escape(this.options.openingTag),
      closingTag: escape(this.options.closingTag),
    };
  }

  isolateTaggedNodes(): this {
    const paragraphs = this.element.getElementsByTagName('a:p');
    const pattern = this.getRegExp();

    for (let p = 0; p < paragraphs.length; p++) {
      const blocks = paragraphs[p].getElementsByTagName('a:r');

      for (let r = 0; r < blocks.length; r++) {
        const block = blocks[r];
        const textContent = this.getTextElement(block).textContent;

        const match = textContent.matchAll(pattern);
        const matches = [...match];

        if (matches.length) {
          this.splitTextBlock(block, matches, textContent);
        }
      }
    }

    // XmlHelper.dump(this.element)

    return this;
  }

  splitTextBlock(
    block: Element,
    matches: RegExpMatchArray[],
    textContent: string,
  ): void {
    const split = this.getCharacterSplit(matches, textContent);

    let lastBlock = block;
    split.forEach((split) => {
      lastBlock = this.insertBlock(lastBlock, split.text);
    });
    block.parentNode.removeChild(block);
  }

  getCharacterSplit(
    matches: RegExpMatchArray[],
    textContent: string,
  ): CharacterSplit[] {
    let lastEnd: number;
    const split = <CharacterSplit[]>[];
    matches.forEach((match, s) => {
      const start = match.index;
      const end = match.index + match[0].length;

      if (s === 0 && start > 0) {
        this.pushCharacterSplit(split, 0, start, textContent);
      }

      if (start > lastEnd) {
        this.pushCharacterSplit(split, lastEnd, match.index, textContent);
      }

      this.pushCharacterSplit(split, start, end, textContent);

      const length = textContent.length;
      if (!matches[s + 1] && end < length) {
        this.pushCharacterSplit(split, end, length, textContent);
      }
      lastEnd = end;
    });
    return split;
  }

  pushCharacterSplit(
    split: CharacterSplit[],
    from: number,
    to: number,
    text: string,
  ): void {
    split.push({
      from: from,
      to: to,
      text: text.slice(from, to),
    });
  }

  insertBlock(block: Element, text: string): Element {
    const newBlock = block.cloneNode(true) as Element;
    const newTextElement = this.getTextElement(newBlock);
    newTextElement.firstChild.textContent = text;

    XmlHelper.insertAfter(newBlock, block);

    return newBlock;
  }

  applyReplacements(replaceTexts: ReplaceText[]): void {
    const textBlocks = this.element.getElementsByTagName('a:r');
    const length = textBlocks.length;

    for (let i = 0; i < length; i++) {
      const textBlock = textBlocks[i];

      replaceTexts.forEach((item) => {
        this.applyReplacement(item, textBlock);
      });
    }
  }

  applyReplacement(replaceText: ReplaceText, textBlock: Element): void {
    const replace =
      this.options.openingTag + replaceText.replace + this.options.closingTag;
    let textNode = this.getTextElement(textBlock);
    const sourceText = textNode.firstChild.textContent;

    if (sourceText.includes(replace)) {
      const bys = GeneralHelper.arrayify(replaceText.by);
      bys.forEach((by, i) => {
        const replacedText = sourceText.replace(replace, by.text);
        textNode = this.assertTextNode(i, textBlock, textNode);
        textNode.firstChild.textContent = replacedText;

        if (by.style) {
          const styleParent = textNode.parentNode as Element;
          const styleElement = styleParent.getElementsByTagName('a:rPr')[0];
          this.applyTextStyle(by.style, styleElement);
        }
      });
    }
  }

  applyTextStyle(style: TextStyle, styleElement: Element): void {
    if (style.color) {
      ModifyTextHelper.setColor(style.color)(styleElement);
    }
    if (style.size) {
      ModifyTextHelper.setSize(style.size)(styleElement);
    }
  }

  assertTextNode(i: number, textBlock: Element, textNode: Element): Element {
    if (i >= 1) {
      const addedTextBlock = textBlock.cloneNode(true) as Element;
      XmlHelper.insertAfter(addedTextBlock, textBlock);
      return this.getTextElement(addedTextBlock);
    }
    return textNode;
  }

  getTextElement(block: Element): Element {
    return block.getElementsByTagName('a:t')[0];
  }

  getRegExp(): RegExp {
    return new RegExp(
      [
        this.expressions.openingTag,
        '[^',
        this.expressions.openingTag,
        this.expressions.closingTag,
        ']+',
        this.expressions.closingTag,
      ].join(''),
      'g',
    );
  }
}
