import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import { GeneralHelper } from './general-helper';
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
    const paragraphs = element.getElementsByTagName('a:p')
    const length = paragraphs.length
    for(let i=0; i<length; i++) {
      const paragraph = paragraphs[i]
      if(i === 0) {
        const blocks = element.getElementsByTagName('a:r')
        const length = blocks.length
        for(let j=0; j<length; j++) {
          const block = blocks[j]
          if(j === 0) {
            const textNode = block.getElementsByTagName('a:t')[0]
            textNode.firstChild.textContent = text;
          } else {
            block.parentNode.removeChild(block);
          }
        }
      } else {
        paragraph.parentNode.removeChild(paragraph);
      }
    }
    // XmlHelper.dump(element)
  };

  /**
   * Replace text content within modified shape
   */
  static replaceText = (replaceText: ReplaceText|ReplaceText[], options?: ReplaceTextOptions) => (element: XMLDocument | Element): void => {
    const replaceTexts = GeneralHelper.arrayify(replaceText);

    new TextReplaceHelper(options, element as XMLDocument)
      .isolateTaggedNodes()
      .applyReplacements(replaceTexts)
  };

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
