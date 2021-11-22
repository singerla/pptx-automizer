import {vd} from './general-helper';
import escape from 'regexp.escape';
import {XmlHelper} from './xml-helper';

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
  newNodes: Element[]

  constructor(options) {
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

    this.matches = <MatchedTag[]>[]
    this.newNodes = <Element[]>[]
  }

  run(textNodes:HTMLCollectionOf<Element>): this {
    const length = textNodes.length
    for(let i=0; i<length; i++) {
      let nodeId = Number(i)

      const textNode = textNodes[i].getElementsByTagName('a:t')[0]
      const text = textNode.firstChild.textContent

      this.testForMatches(text, 'openingTag', nodeId)
      this.testForMatches(text, 'closingTag', nodeId)
    }

    this.sortMatches()
    this.cloneNodes(textNodes)

    return this
  }

  replaceChildren(parent: Element) {
    if(this.newNodes.length === 0) return

    const blocks = parent.getElementsByTagName('a:r')
    const length = blocks.length
    for(let i=0; i<length; i++) {
      const block = blocks[i]
      block.parentNode.removeChild(block);
    }

    this.newNodes.forEach(newNode => {
      parent.appendChild(newNode)
    })

    return parent
  }

  cloneNodes = (textNodes) => {
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

  getTextNode(block:Element) {
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

  pushMatchedNode = (sourceNodes: Element[], openingMatch:MatchedTag, endingMatch:MatchedTag): void => {
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
        tagParts.push(sourceText.slice(openingMatch.index, endingMatch.index+1))
      } else {
        tagParts.push(sourceText)
      }
    }

    newNode.textContent = tagParts.join('')

    this.newNodes.push(newBlock)
  }

  stripPrependingTextTag(text:string, match:MatchedTag) {
    return text.slice(0, match.index - 1)
  }
  stripAppendingTextTag(text, match:MatchedTag) {
    return text.slice(match.index + 1)
  }

  stripPrependingText(text:string, match:MatchedTag) {
    return text.slice(match.index)
  }
  stripAppendingText(text, match:MatchedTag) {
    return text.slice(0, match.index + 1)
  }

  testForMatches = (text:string, type:MatchType, nodeId:number) => {
    const regExp = new RegExp(this.expressions[type].escaped, 'g')
    const allMatches = [...text.matchAll(regExp)]
    this.parseMatches(allMatches, type, nodeId, text)
  }

  parseMatches = (matchedAll:RegExpMatchArray[], type:MatchType, nodeId:number, text:string) => {
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
      const remainingText = match.input.slice(0, match.index)
      if(remainingText.match(this.expressions.openingTag.escaped)) {
        return 0
      }
      return index - 1
    }

    if(type === 'closingTag' && length > index) {
      const remainingText = match.input.slice(match.index+1)
      if(remainingText.match(this.expressions.openingTag.escaped)) {
        return 0
      }
      return length - index - 1
    }
    return 0
  }

  sortMatches() {
    this.matches.sort((a, b) => {
      return a.nodeId - b.nodeId || a.index - b.index
    })
  }
}
