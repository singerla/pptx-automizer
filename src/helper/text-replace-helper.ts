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

    for(let p=0; p<length; p++) {
      const paragraph = paragraphs[p]
      const textBlocks = paragraph.getElementsByTagName('a:r')

      if(textBlocks.length === 0) continue

      this.getSortedTextBlocks(textBlocks)
        .replaceChildren(paragraphs[p])
    }

    // XmlHelper.dump(this.element)
    
    return this
  }

  applyReplacements(replaceTexts:ReplaceText[]): void {
    const textBlocks = this.element.getElementsByTagName('a:r')
    const length = textBlocks.length

    for(let i=0; i<length; i++) {
      const textBlock = textBlocks[i]

      replaceTexts.forEach(item => {
        const replace = this.options.openingTag + item.replace + this.options.closingTag
        let textNode = textBlock.getElementsByTagName('a:t')[0]
        const sourceText = textNode.firstChild.textContent
        const match = sourceText.includes(replace)
        const bys = GeneralHelper.arrayify(item.by);

        if(match === true) {
          bys.forEach((by,i) => {
            const replacedText = sourceText.replace(replace, by.text)
            if(i >= 1) {
              const addedTextBlock = textBlock.cloneNode(true) as Element
              XmlHelper.insertAfter(addedTextBlock, textBlock)
              textNode = addedTextBlock.getElementsByTagName('a:t')[0]
            }
            textNode.firstChild.textContent = replacedText
          })
        }
      })
    }
  }

  getSortedTextBlocks(textBlocks:HTMLCollectionOf<Element>): this {
    this.matches = <MatchedTag[]>[]
    this.newNodes = <Element[]>[]

    const length = textBlocks.length
    for(let i=0; i<length; i++) {
      const nodeId = Number(i)

      const textNode = textBlocks[i].getElementsByTagName('a:t')[0]
      const text = textNode.firstChild.textContent

      this.testForMatches(text, 'openingTag', nodeId)
      this.testForMatches(text, 'closingTag', nodeId)
    }

    this.sortMatches()
    this.cloneNodes(textBlocks)

    return this
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

  cloneNodes = (textNodes: Element[]|HTMLCollectionOf<Element>): void => {
    let lastNode = 0
    this.matches.forEach((match,m) => {
      if(match.type === 'openingTag') {
        if(match.nodeId > lastNode) {
          for(let i=lastNode; i<match.nodeId; i++) {
            this.pushDefaultNode(textNodes[i], 'addPrependingText')
          }
        }

        if(match.remaining > 0) {
          this.pushDefaultNode(textNodes[match.nodeId], 'addPrependingText', match)
        }

        const endingMatch = this.matches[m+1]
        this.pushMatchedNode(textNodes, match, endingMatch)

        lastNode = endingMatch.nodeId + 1
      }

      if(match.type === 'closingTag' && match.remaining > 0) {
        this.pushDefaultNode(textNodes[match.nodeId], 'addAppendingText', match)
      }
    })
  }

  getTextNode(block:Element): Element {
    return block.getElementsByTagName('a:t')[0]
  }

  pushDefaultNode = (sourceNode: Element, mode:string, match?:MatchedTag): void => {
    const newBlock = sourceNode.cloneNode(true) as Element
    const textNode = this.getTextNode(newBlock)

    if(match) {
      switch (mode) {
        case 'addPrependingText':
          textNode.textContent = this.stripPrependingTextTag(textNode.textContent, match)
          break;
        case 'addAppendingText':
          textNode.textContent = this.stripAppendingTextTag(textNode.textContent, match)
          break;
      }
    }
    this.newNodes.push(newBlock)
  }

  pushMatchedNode = (sourceNodes: Element[]|HTMLCollectionOf<Element>, openingMatch:MatchedTag, endingMatch:MatchedTag): void => {
    const sourceNodeId = openingMatch.nodeId
    const newBlock = sourceNodes[sourceNodeId].cloneNode(true) as Element
    const newNode = this.getTextNode(newBlock)

    const tagParts = <string[]>[]
    for(let i=openingMatch.nodeId; i<=endingMatch.nodeId; i++) {
      const sourceText = sourceNodes[i].textContent

      if(i === openingMatch.nodeId && openingMatch.remaining > 0) {
        tagParts.push(this.stripPrependingText(sourceText, openingMatch))
      } else if(i === endingMatch.nodeId && endingMatch.remaining > 0) {
        tagParts.push(this.stripAppendingText(sourceText, endingMatch))
      } else if(i === openingMatch.nodeId && openingMatch.nodeId === endingMatch.nodeId) {
        const length = endingMatch.index + this.expressions.closingTag.length
        tagParts.push(sourceText.slice(openingMatch.index, length))
      } else {
        tagParts.push(sourceText)
      }
    }

    newNode.textContent = tagParts.join('')

    this.newNodes.push(newBlock)
  }

  stripPrependingTextTag(text:string, match:MatchedTag): string {
    return text.slice(0, match.index)
  }
  stripAppendingTextTag(text:string, match:MatchedTag): string {
    return text.slice(match.index + this.options.closingTag.length)
  }
  stripPrependingText(text:string, match:MatchedTag): string {
    return text.slice(match.index)
  }
  stripAppendingText(text:string, match:MatchedTag): string {
    return text.slice(0, match.index + this.options.closingTag.length)
  }

  testForMatches = (text:string, type:MatchType, nodeId:number): void => {
    const regExp = new RegExp(this.expressions[type].escaped, 'g')
    const allMatches = [...text.matchAll(regExp)]
    this.parseMatches(allMatches, type, nodeId, text)
  }

  parseMatches = (matchedAll:RegExpMatchArray[], type:MatchType, nodeId:number, text:string): void => {
    matchedAll.forEach(match => {
      if(match.length > 0) {
        this.matches.push({
          type: type,
          index: match.index,
          nodeId: nodeId,
          match: match,
          text: text,
          remaining: this.checkRemainingChars(type, match)
        })
      }
    })
  }

  checkRemainingChars(type:MatchType, match:RegExpMatchArray): number {
    const length = match.input.length
    const index = match.index

    if(type === 'openingTag' && index > 0) {
      // const remainingText = match.input.slice(0, match.index)
      // const hasClosingTag = remainingText.match(this.expressions.closingTag.escaped)
      // if(hasClosingTag) {
      //   return remainingText.length - hasClosingTag.index
      // }
      return index - 1
    }

    if(type === 'closingTag' && length > index) {
      const closingTagLength = this.options.closingTag.length
      // const remainingText = match.input.slice(match.index + closingTagLength)
      // const hasOpeningTag = remainingText.match(this.expressions.openingTag.escaped)
      // if(hasOpeningTag) {
      //   return 0
      // }
      return length - index - closingTagLength
    }
    return 0
  }

  sortMatches(): void {
    this.matches.sort((a, b) => {
      return a.nodeId - b.nodeId || a.index - b.index
    })
  }
}
