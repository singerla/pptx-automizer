import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import {XmlHelper} from './xml-helper';
import {GeneralHelper, vd} from './general-helper';
import TextReplaceHelper from './text-replace-helper';

export default class ModifyShapeHelper {
  /**
   * Set solid fill of modified shape
   */
  static setSolidFill = (element: XMLDocument | Element): void => {
    element
      .getElementsByTagName('a:solidFill')[0]
      .getElementsByTagName('a:schemeClr')[0]
      .setAttribute('val', 'accent6');
  };

  /**
   * Set text content of modified shape
   */
  static setText = (text: string) => (element: XMLDocument | Element): void => {
    const textNodes = element.getElementsByTagName('a:t')
    textNodes[0].firstChild.textContent = text;
    // TODO: get rid of remaining text nodes
    // XmlHelper.sliceCollection(textNodes, textNodes.length-1)
    // XmlHelper.dump(element)
  };

  /**
   * Replace text content within modified shape
   */
  static replaceText = (replaceText: ReplaceText|ReplaceText[], options?: ReplaceTextOptions) => (element: XMLDocument | Element): void => {
    const replaceTexts = GeneralHelper.arrayify(replaceText);
    const defaultOptions = {
      openingTag: '${',
      closingTag: '}'
    }
    options = (!options) ? defaultOptions : {...defaultOptions, ...options}

    ModifyShapeHelper.isolateTags(element, options)

    const textBlocks = element.getElementsByTagName('a:r')
    const length = textBlocks.length
    for(let i=0; i<length; i++) {
      const textBlock = textBlocks[i]

      replaceTexts.forEach(item => {
        const replace = defaultOptions.openingTag + item.replace + defaultOptions.closingTag
        const textNode = textBlock.getElementsByTagName('a:t')[0]
        const sourceText = textNode.firstChild.textContent
        const match = sourceText.includes(replace)
        const bys = GeneralHelper.arrayify(item.by);

        if(match === true) {
          // TODO: clone original text block to add more than one replacement
          // if(bys.length > 1) {
          //   for(i=1;i<=bys.length;i++) {
          //     textBlock.parentNode.appendChild(textBlock.cloneNode(true))
          //   }
          // }

          bys.forEach((by,i) => {
            textNode.firstChild.textContent = textNode.firstChild.textContent.replace(replace, by.text)
          })
        }
      })
    }
  };

  /**
   * Assert all tags to be isolated (separated) into a single <a:r>-element.
   * This is required to have a clean source object for possibly cloned text nodes.
   * @param element
   */
  static isolateTags = (element: XMLDocument | Element, options: ReplaceTextOptions): void => {
    const paragraphs = element.getElementsByTagName('a:p')
    const length = paragraphs.length

    for(let p=0; p<length; p++) {
      const paragraph = paragraphs[p]
      const textBlocks = paragraph.getElementsByTagName('a:r')

      if(textBlocks.length === 0) continue

      const textReplacementHelper = new TextReplaceHelper(options)
      textReplacementHelper
        .run(textBlocks)
        .replaceChildren(paragraphs[p])
    }
  }

  /**
   * Set position and size of modified shape.
   */
  static setPosition = (pos: ShapeCoordinates) => (
    element: XMLDocument | Element,
  ): void => {
    const map = {
      x: { tag: 'a:off', attribute: 'x' },
      y: { tag: 'a:off', attribute: 'y' },
      w: { tag: 'a:ext', attribute: 'cx' },
      h: { tag: 'a:ext', attribute: 'cy' },
      cx: { tag: 'a:ext', attribute: 'cx' },
      cy: { tag: 'a:ext', attribute: 'cy' },
    };

    const xfrm = element.getElementsByTagName('a:off')[0].parentNode as Element

    Object.keys(pos).forEach((key) => {
      xfrm.getElementsByTagName(map[key].tag)[0]
        .setAttribute(map[key].attribute, Math.round(pos[key]));
    });
  };
}
