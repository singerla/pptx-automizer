import {Color} from '../types/modify-types';
import { vd } from './general-helper';
import XmlElements from './xml-elements';
import {XmlHelper} from './xml-helper';

export default class ModifyColorHelper {
  /**
   * Replaces or creates an <a:solidFill> Element
   */
  static solidFill = (color: Color) => (element: Element): void => {
    const solidFills = element.getElementsByTagName('a:solidFill')

    if(!solidFills.length) {
      const solidFill = new XmlElements(element, {
        color: color
      }).solidFill()
      element.appendChild(solidFill)
      return
    }

    const solidFill = solidFills[0] as Element
    const colorType = new XmlElements(element, {
      color: color
    }).colorType()

    XmlHelper.sliceCollection(solidFill.childNodes as unknown as HTMLCollectionOf<Element>, 0)
    solidFill.appendChild(colorType)
  }
}
