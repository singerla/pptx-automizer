import { Color } from '../types/modify-types';
import { vd } from './general-helper';
import XmlElements from './xml-elements';
import { XmlHelper } from './xml-helper';

export default class ModifyColorHelper {
  /**
   * Replaces or creates an <a:solidFill> Element
   */
  static solidFill =
    (color: Color, index?: number | 'last') =>
    (element: Element): void => {
      if (!color || !color.type) return;

      const solidFills = element.getElementsByTagName('a:solidFill');

      if (!solidFills.length) {
        const solidFill = new XmlElements(element, {
          color: color,
        }).solidFill();
        element.appendChild(solidFill);
        return;
      }

      let targetIndex = !index
        ? 0
        : index === 'last'
        ? solidFills.length - 1
        : index;

      const solidFill = solidFills[targetIndex] as Element;
      const colorType = new XmlElements(element, {
        color: color,
      }).colorType();

      XmlHelper.sliceCollection(
        solidFill.childNodes as unknown as HTMLCollectionOf<Element>,
        0,
      );
      solidFill.appendChild(colorType);
    };

  static removeNoFill =
    () =>
    (element: Element): void => {
      const hasNoFill = element.getElementsByTagName('a:noFill')[0];
      if (hasNoFill) {
        element.removeChild(hasNoFill);
      }
    };
}
