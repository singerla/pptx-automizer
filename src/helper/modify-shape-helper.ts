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

  /**
   * Apply rounded corners to a shape with a fixed corner radius
   * @param degree Corner radius in EMU units (1 cm = 360000 EMU)
   */
  static roundedCorners = (degree: number) => (element: XmlElement): void => {
    // Find the spPr element where we need to add or modify the a:prstGeom element
    const spPr = element.getElementsByTagName('p:spPr')[0] ||
      element.getElementsByTagName('a:spPr')[0];

    if (!spPr) {
      return; // Cannot find spPr element
    }

    // Get the shape dimensions to calculate the appropriate adjustment value
    const xfrm = ModifyShapeHelper.ensureTransformElements(element);
    if (!xfrm) {
      return; // Cannot proceed without proper transformation data
    }

    // Get current width and height
    const width = Number(xfrm.getElementsByTagName('a:ext')[0].getAttribute('cx'));
    const height = Number(xfrm.getElementsByTagName('a:ext')[0].getAttribute('cy'));

    // Calculate the adjustment value (percentage)
    // The adjustment is a percentage (0-100000) of the smaller dimension
    const minDimension = Math.min(width, height);

    // Ensure degree is within reasonable bounds (PowerPoint uses 0-50% for rounded rect)
    const clampedDegree = Math.max(0, Math.min(degree, minDimension / 2));

    // Calculate the adjustment value (0-100000 where 100000 is 100%)
    // PowerPoint uses the percentage of the shorter dimension for corners
    const adjValue = Math.round((clampedDegree / minDimension) * 100000);

    // Remove any existing prstGeom element
    const existingPrstGeom = spPr.getElementsByTagName('a:prstGeom')[0];
    if (existingPrstGeom) {
      spPr.removeChild(existingPrstGeom);
    }

    // Create the new prstGeom element with the roundRect preset
    const prstGeom = element.ownerDocument.createElement('a:prstGeom');
    prstGeom.setAttribute('prst', 'roundRect');

    // Create the avLst element and the adjustment value
    const avLst = element.ownerDocument.createElement('a:avLst');
    const gd = element.ownerDocument.createElement('a:gd');
    gd.setAttribute('name', 'adj');
    gd.setAttribute('fmla', `val ${adjValue}`);

    // Build the element hierarchy
    avLst.appendChild(gd);
    prstGeom.appendChild(avLst);

    // Add the new prstGeom element to spPr at the appropriate position
    // It should be added after xfrm but before other elements
    const xfrmElement = spPr.getElementsByTagName('a:xfrm')[0];

    if (xfrmElement && xfrmElement.nextSibling) {
      spPr.insertBefore(prstGeom, xfrmElement.nextSibling);
    } else {
      spPr.appendChild(prstGeom);
    }

    // Check for noFill element - for picture elements, we need to remove noFill
    // otherwise the picture will be invisible when we apply rounded corners
    const noFillElement = spPr.getElementsByTagName('a:noFill')[0];
    if (noFillElement) {
      spPr.removeChild(noFillElement);
    }
  }
}
