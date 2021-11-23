import {GeneralHelper, vd} from './general-helper';
import escape from 'regexp.escape';
import {XmlHelper} from './xml-helper';
import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';

type MatchType = 'openingTag'|'closingTag'
type MatchedTag = {
  type: MatchType;
  index: number;
  nodeId: number;
  text: string;
  remaining: number;
  match: RegExpMatchArray;
}
type Expressions = {
  [key in MatchType]: {
    escaped: string;
    length: number;
  };
};
type MatchedFragment = {
  type: 'tag'|'text';
  text: string;
  from: number;
  to: number;
}

export default class TextReplaceHelper {
  expressions: Expressions
  matches: MatchedTag[]
  element: XMLDocument
  newNodes: Element[]
  options: ReplaceTextOptions

  constructor(options: ReplaceTextOptions, element:XMLDocument) {
    const defaultOptions = {
      openingTag: '{{',
      closingTag: '}}'
    }

    this.options = (!options)
      ? defaultOptions
      : {...defaultOptions, ...options}

    this.element = element
    this.expressions = {
      openingTag: {
        escaped: escape(options.openingTag),
        length: options.openingTag.length
      },
      closingTag: {
        escaped: escape(options.closingTag),
        length: options.closingTag.length
      }
    }
  }

  isolateTaggedNodes(): this {
    const paragraphs = this.element.getElementsByTagName('a:p')
    const length = paragraphs.length

    // XmlHelper.dump(this.element)

    for(let p=0; p<length; p++) {
      const paragraph = paragraphs[p]
      const textBlocks = paragraph.getElementsByTagName('a:r')

      if(textBlocks.length === 0) continue

      this.getSortedTextBlocks(textBlocks)
        .replaceChildren(paragraphs[p])
    }

    return this
  }

  applyReplacements(replaceTexts:ReplaceText[]): void {
    const textBlocks = this.element.getElementsByTagName('a:r')
    const length = textBlocks.length

    for(let i=0; i<length; i++) {
      const textBlock = textBlocks[i]

      replaceTexts.forEach(item => {
        this.applyReplacement(item, textBlock)
      })
    }
  }

  applyReplacement(replaceText: ReplaceText, textBlock: Element): void {
    const replace = this.options.openingTag + replaceText.replace + this.options.closingTag
    let textNode = textBlock.getElementsByTagName('a:t')[0]

    const sourceText = textNode.firstChild.textContent
    if(sourceText.includes(replace)) {
      const bys = GeneralHelper.arrayify(replaceText.by);
      bys.forEach((by,i) => {
        const replacedText = sourceText.replace(replace, by.text)
        textNode = this.assertTextNode(i, textBlock, textNode)
        textNode.firstChild.textContent = replacedText
      })
    }
  }

  assertTextNode(i:number, textBlock:Element, textNode:Element): Element {
    if(i >= 1) {
      const addedTextBlock = textBlock.cloneNode(true) as Element
      XmlHelper.insertAfter(addedTextBlock, textBlock)
      return addedTextBlock.getElementsByTagName('a:t')[0]
    }
    return textNode
  }

  getSortedTextBlocks(textBlocks:HTMLCollectionOf<Element>): this {
    this.matches = <MatchedTag[]>[]
    this.newNodes = <Element[]>[]

    const length = textBlocks.length
    const fragments = []
    const mapBlocks = []

    let currentLength = 0
    for(let i=0; i<length; i++) {
      const nodeId = Number(i)

      const textNode = textBlocks[i].getElementsByTagName('a:t')[0]
      const text = textNode.firstChild.textContent

      mapBlocks.push(
        {
          text: text,
          nodeId: nodeId,
          from: currentLength,
          to: currentLength + text.length,
          node: textBlocks[i]
        }
      )

      fragments.push(text)

      currentLength += text.length
    }

    const fullText = fragments.join('')
    const pattern = this.getRegExpPattern()
    const regExp = new RegExp(pattern, 'g')
    const matchAll = fullText.matchAll(regExp)
    const allMatches = [...matchAll]

    const matchedFragments = this.getMatchedFragments(allMatches, fullText)
    matchedFragments.forEach(fragment => {
      if(fragment.type === 'text') {
        const blocksToPush = mapBlocks.filter(mapBlock => mapBlock.from >= fragment.from && mapBlock.to <= fragment.to)
        if(blocksToPush.length === 0) {
          const blockToPush = mapBlocks.find(mapBlock => fragment.from >= mapBlock.from)
          this.pushNewNode(blockToPush.node, fragment.text)
        } else {
          blocksToPush.forEach(blockToPush => {
            this.pushNewNode(blockToPush.node, blockToPush.text)
          })
        }
      } else {
        const blockToPush = mapBlocks.find(mapBlock => mapBlock.to >= fragment.to)
        this.pushNewNode(blockToPush.node, fragment.text)
      }
    })

    return this
  }

  pushNewNode = (sourceNode: Element, textContent?:string): void => {
    const newBlock = sourceNode.cloneNode(true) as Element

    if(textContent) {
      const textNode = newBlock.getElementsByTagName('a:t')[0]
      textNode.textContent = textContent
    }

    this.newNodes.push(newBlock)
  }

  getMatchedFragments(allMatches: RegExpMatchArray[], fullText:string): MatchedFragment[] {
    const matchedFragments = <MatchedFragment[]>[]
    if(allMatches.length > 0) {
      let lastIndex = 0
      let currentIndex = 0
      allMatches.forEach((match, m) => {
        if(match.index > lastIndex) {
          this.pushFragment('text', fullText.slice(lastIndex, match.index), matchedFragments, lastIndex)
        }
        this.pushFragment('tag', fullText.slice(match.index, match.index + match[0].length), matchedFragments, match.index)
        lastIndex = match.index + match[0].length
        if(!allMatches[m+1] && fullText.length > lastIndex) {
          this.pushFragment('text', fullText.slice(lastIndex), matchedFragments, lastIndex)
        }
        currentIndex += lastIndex
      })
    } else {
      matchedFragments.push({
        type: 'text',
        text: fullText,
        from: 0,
        to: fullText.length
      })
    }
    return matchedFragments
  }

  pushFragment(type:MatchedFragment['type'], text:string, matchedFragments:MatchedFragment[], lastIndex:number) {
    matchedFragments.push({
      type: type,
      text: text,
      from: lastIndex,
      to: lastIndex + text.length
    })
  }

  getRegExpPattern(): string {
    return [
      this.expressions.openingTag.escaped,
      '[^',
      this.expressions.openingTag.escaped,
      this.expressions.closingTag.escaped,
      ']+',
      this.expressions.closingTag.escaped,
    ].join('')
  }

  replaceChildren(paragraph: Element): Element {
    if(this.newNodes.length === 0) return

    const blocks = paragraph.getElementsByTagName('a:r')
    const length = blocks.length
    for(let i=0; i<length; i++) {
      const block = blocks[i]
      block.parentNode.removeChild(block);
    }

    this.newNodes.forEach(newNode => {
      paragraph.appendChild(newNode)
    })

    return paragraph
  }
}
