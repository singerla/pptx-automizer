import { Color } from '../types/modify-types';
import XmlElements from './xml-elements';
import {vd} from './general-helper';
import {XmlHelper} from './xml-helper';

export default class ModifyTextHelper {
  /**
   * Set color of text insinde an <a:rPr> element
   */
  static setColor = (color: Color) => (element: Element): void => {
    if(!color) return

    new XmlElements(element, {
      color: color,
    }).solidFill();
  };

  /**
   * Set size of text insinde an <a:rPr> element
   */
  static setSize = (size: number) => (element: Element): void => {
    element.setAttribute('sz', String(Math.round(size)));
  };
}
