import { ReplaceText } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import {XmlHelper} from './xml-helper';

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
  static replaceText = (replaceText: ReplaceText[]) => (element: XMLDocument | Element): void => {
    const textNodes = element.getElementsByTagName('a:t')
    for(const i in textNodes) {
      if(!textNodes[i].firstChild?.textContent) continue
      const textContent = textNodes[i].firstChild.textContent
      replaceText.forEach(item => {
        const replacedContent = textContent.replace(item.replace, item.by)
        textNodes[i].firstChild.textContent = replacedContent
      })
    }
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
