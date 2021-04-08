import { ShapeCoordinates } from '../types/shape-types';

export default class ModifyShapeHelper {
  /**
   * Set solid fill of modified shape
   */
  static setSolidFill = (element: XMLDocument): void => {
    element
      .getElementsByTagName('a:solidFill')[0]
      .getElementsByTagName('a:schemeClr')[0]
      .setAttribute('val', 'accent6');
  };

  /**
   * Set text content of modified shape
   */
  static setText = (text: string) => (element: XMLDocument): void => {
    element.getElementsByTagName('a:t')[0].firstChild.textContent = text;
  };

  /**
   * Set position and size of modified shape.
   */
  static setPosition = (pos: ShapeCoordinates) => (
    element: XMLDocument,
  ): void => {
    const map = {
      x: { tag: 'a:off', attribute: 'x' },
      y: { tag: 'a:off', attribute: 'y' },
      w: { tag: 'a:ext', attribute: 'cx' },
      h: { tag: 'a:ext', attribute: 'cy' },
    };

    const parent = 'a:xfrm';

    Object.keys(pos).forEach((key) => {
      element
        .getElementsByTagName(parent)[0]
        .getElementsByTagName(map[key].tag)[0]
        .setAttribute(map[key].attribute, pos[key]);
    });
  };
}
