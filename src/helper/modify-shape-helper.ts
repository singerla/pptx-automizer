import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import { GeneralHelper } from './general-helper';
import TextReplaceHelper from './text-replace-helper';
import ModifyTextHelper from './modify-text-helper';
import { XmlElement } from '../types/xml-types';

const map = {
  x: { tag: 'a:off', attribute: 'x' },
  l: { tag: 'a:off', attribute: 'x' },
  left: { tag: 'a:off', attribute: 'x' },
  y: { tag: 'a:off', attribute: 'y' },
  t: { tag: 'a:off', attribute: 'y' },
  top: { tag: 'a:off', attribute: 'y' },
  cx: { tag: 'a:ext', attribute: 'cx' },
  w: { tag: 'a:ext', attribute: 'cx' },
  width: { tag: 'a:ext', attribute: 'cx' },
  cy: { tag: 'a:ext', attribute: 'cy' },
  h: { tag: 'a:ext', attribute: 'cy' },
  height: { tag: 'a:ext', attribute: 'cy' },
};

export default class ModifyShapeHelper {
  /**
   * Set solid fill of modified shape
   */
  static setSolidFill = (element: XmlElement): void => {
    element
      .getElementsByTagName('a:solidFill')[0]
      .getElementsByTagName('a:schemeClr')[0]
      .setAttribute('val', 'accent6');
  };

  /**
   * Set text content of modified shape
   */
  static setText =
    (text: string) =>
    (element: XmlElement): void => {
      ModifyTextHelper.setText(text)(element as XmlElement);
    };

  /**
   * Set content to bulleted list of modified shape
   */
  static setBulletList =
    (list) =>
    (element: XmlElement): void => {
      ModifyTextHelper.setBulletList(list)(element as XmlElement);
    };

  /**
   * Replace tagged text content within modified shape
   */
  static replaceText =
    (replaceText: ReplaceText | ReplaceText[], options?: ReplaceTextOptions) =>
    (element: XmlElement): void => {
      const replaceTexts = GeneralHelper.arrayify(replaceText);

      new TextReplaceHelper(options, element as XmlElement)
        .isolateTaggedNodes()
        .applyReplacements(replaceTexts);
    };

  /**
   * Set position and size of modified shape.
   */
  static setPosition =
    (pos: ShapeCoordinates) =>
    (element: XmlElement): void => {
      const aOff = element.getElementsByTagName('a:off');
      if (!aOff?.item(0)) {
        return;
      }

      const xfrm = aOff.item(0).parentNode as XmlElement;

      Object.keys(pos).forEach((key) => {
        let value = Math.round(pos[key]);
        if (typeof value !== 'number' || !map[key]) return;
        value = value < 0 ? 0 : value;

        xfrm
          .getElementsByTagName(map[key].tag)[0]
          .setAttribute(map[key].attribute, value);
      });
    };

  /**
   * Update position and size of a shape by a given Value.
   */
  static updatePosition =
    (pos: ShapeCoordinates) =>
    (element: XmlElement): void => {
      const xfrm = element.getElementsByTagName('a:off')[0]
        .parentNode as XmlElement;

      Object.keys(pos).forEach((key) => {
        let value = Math.round(pos[key]);
        if (typeof value !== 'number' || !map[key]) return;

        const currentValue = xfrm
          .getElementsByTagName(map[key].tag)[0]
          .getAttribute(map[key].attribute);

        value += Number(currentValue);

        xfrm
          .getElementsByTagName(map[key].tag)[0]
          .setAttribute(map[key].attribute, value);
      });
    };

  /**
   * Rotate a shape by a given value. Use e.g. 180 to flip a shape.
   * A negative value will rotate counter clockwise.
   * @param degrees Rotate by °
   */
  static rotate =
    (degrees: number) =>
    (element: XmlElement): void => {
      const spPr = element.getElementsByTagName('p:spPr');

      if (spPr) {
        const xfrm = spPr.item(0).getElementsByTagName('a:xfrm').item(0);
        degrees = degrees < 0 ? 360 + degrees : degrees;
        xfrm.setAttribute('rot', String(Math.round(degrees * 60000)));
      }
    };
}
