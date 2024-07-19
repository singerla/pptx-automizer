import { Color, ImageStyle } from '../types/modify-types';
import XmlElements from './xml-elements';
import { XmlHelper } from './xml-helper';
import { XmlElement } from '../types/xml-types';

export default class ModifyColorHelper {
  /**
   * Replaces or creates an <a:solidFill> Element
   */
  static solidFill =
    (color: Color, index?: number | 'last') =>
    (element: XmlElement): void => {
      if (!color || !color.type || element?.getElementsByTagName === undefined)
        return;

      ModifyColorHelper.normalizeColorObject(color);

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

      const solidFill = solidFills[targetIndex] as XmlElement;
      const colorType = new XmlElements(element, {
        color: color,
      }).colorType();

      XmlHelper.sliceCollection(
        solidFill.childNodes as unknown as HTMLCollectionOf<XmlElement>,
        0,
      );
      solidFill.appendChild(colorType);
    };

  static removeNoFill =
    () =>
    (element: XmlElement): void => {
      const hasNoFill = element.getElementsByTagName('a:noFill')[0];
      if (hasNoFill) {
        element.removeChild(hasNoFill);
      }
    };

  static normalizeColorObject = (color: Color) => {
    if (color.value.indexOf('#') === 0) {
      color.value = color.value.replace('#', '');
    }
    if (color.value.toLowerCase().indexOf('rgb(') === 0) {
      // TODO: convert RGB to HEX
      color.value = 'cccccc';
    }
    return color;
  };
}
