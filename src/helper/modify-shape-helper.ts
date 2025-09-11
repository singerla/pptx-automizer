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
   * Creates missing transformation elements (a:off and a:ext) with default values
   * @param element The XML element to check and modify
   * @returns The xfrm element that contains or will contain the a:off element
   */
  static ensureTransformElements = (element: XmlElement): XmlElement | null => {
    // First find the xfrm element (could be a:xfrm or p:xfrm)
    let xfrm = element.getElementsByTagName('a:xfrm')[0] as XmlElement;
    if (!xfrm) {
      xfrm = element.getElementsByTagName('p:xfrm')[0] as XmlElement;
    }

    // If no xfrm element exists, try to find the appropriate parent to create it
    if (!xfrm) {
      const spPr =
        element.getElementsByTagName('p:spPr')[0] ||
        element.getElementsByTagName('a:spPr')[0];

      if (!spPr) {
        return null; // Cannot create xfrm without a proper parent
      }

      // Create the xfrm element
      xfrm = element.ownerDocument.createElement('a:xfrm');
      spPr.appendChild(xfrm);
    }

    // Create a:off element if it doesn't exist
    if (!xfrm.getElementsByTagName('a:off')[0]) {
      const newAOff = element.ownerDocument.createElement('a:off');
      // Set default coordinates
      newAOff.setAttribute('x', '0');
      newAOff.setAttribute('y', '0');

      // Insert a:off as the first child of xfrm
      if (xfrm.hasChildNodes()) {
        xfrm.insertBefore(newAOff, xfrm.firstChild);
      } else {
        xfrm.appendChild(newAOff);
      }
    }

    // Create a:ext element if it doesn't exist
    if (!xfrm.getElementsByTagName('a:ext')[0]) {
      const newAExt = element.ownerDocument.createElement('a:ext');
      // Set default dimensions - using 1000000 which is about 2.78cm
      newAExt.setAttribute('cx', '1000000');
      newAExt.setAttribute('cy', '1000000');

      // Add the a:ext element after a:off
      xfrm.appendChild(newAExt);
    }

    return xfrm;
  };

  /**
   * Set position and size of modified shape.
   */
  static setPosition =
    (pos: ShapeCoordinates) =>
    (element: XmlElement): void => {
      // Ensure the transform elements exist
      const xfrm = ModifyShapeHelper.ensureTransformElements(element);
      if (!xfrm) {
        return; // Cannot set position without proper structure
      }

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
      // Ensure the transform elements exist
      const xfrm = ModifyShapeHelper.ensureTransformElements(element);
      if (!xfrm) {
        return; // Cannot update position without proper structure
      }

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
