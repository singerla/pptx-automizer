import { GeneralHelper, vd } from './general-helper';
import escape from 'regexp.escape';
import { XmlHelper } from './xml-helper';
import {
  ReplaceText,
  ReplaceTextOptions,
  TextStyle,
} from '../types/modify-types';
import ModifyTextHelper from './modify-text-helper';
import { XmlDocument, XmlElement } from '../types/xml-types';
import XmlElements from './xml-elements';

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
  element: XmlElement;
  newNodes: XmlElement[];
  options: ReplaceTextOptions;

  constructor(options: ReplaceTextOptions, element: XmlElement) {
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
    block: XmlElement,
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

  insertBlock(block: XmlElement, text: string): XmlElement {
    const newBlock = block.cloneNode(true) as XmlElement;
    const newTextElement = this.getTextElement(newBlock);
    ModifyTextHelper.content(text)(newTextElement);

    XmlHelper.insertAfter(newBlock, block);

    return newBlock;
  }

  applyReplacements(replaceTexts: ReplaceText[]): void {
    const textBlocks = this.element.getElementsByTagName('a:r');
    const length = textBlocks.length;

    for (let i = 0; i < length; i++) {
      const textBlock = textBlocks[i];

      replaceTexts.forEach((item) => {
        this.applyReplacement(item, textBlock, i);
      });
    }
  }

  applyReplacement(
    replaceText: ReplaceText,
    textBlock: XmlElement,
    currentIndex: number,
  ): void {
    const replace =
      this.options.openingTag + replaceText.replace + this.options.closingTag;
    let textNode = this.getTextElement(textBlock);
    const sourceText = textNode.firstChild?.textContent;

    if (sourceText?.includes(replace)) {
      const bys = GeneralHelper.arrayify(replaceText.by);
      const modifyBlocks = this.assertTextBlocks(bys.length, textBlock);

      bys.forEach((by, blockIndex) => {
        const textNode =
          modifyBlocks[blockIndex].getElementsByTagName('a:t')[0];
        this.updateTextNode(textNode, sourceText, replace, by);
      });
    }
  }

  assertTextBlocks(length: number, textBlock: any): XmlElement[] {
    const modifyBlocks = [];
    if (length > 1) {
      for (let i = 1; i < length; i++) {
        const addedTextBlock = textBlock.cloneNode(true) as XmlElement;
        XmlHelper.insertAfter(addedTextBlock, textBlock);
        modifyBlocks.push(addedTextBlock);
      }
    }
    modifyBlocks.push(textBlock);
    modifyBlocks.reverse();
    return modifyBlocks;
  }

  updateTextNode(textNode: XmlElement, sourceText, replace, by): void {
    const replacedText = sourceText.replace(replace, by.text);
    ModifyTextHelper.content(replacedText)(textNode);

    if (by.style) {
      const styleParent = textNode.parentNode as XmlElement;
      const styleElement = styleParent.getElementsByTagName('a:rPr')[0];
      ModifyTextHelper.style(by.style)(styleElement);
    }
  }

  getTextElement(block: XmlElement): XmlElement {
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
