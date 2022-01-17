import { ReplaceText, ReplaceTextOptions } from '../types/modify-types';
import { ShapeCoordinates } from '../types/shape-types';
import { GeneralHelper } from './general-helper';
import TextReplaceHelper from './text-replace-helper';
import ModifyTextHelper from './modify-text-helper';

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
    ModifyTextHelper.setText(text)(element as Element)
  };

  /**
   * Replace tagged text content within modified shape
   */
  static replaceText = (
    replaceText: ReplaceText | ReplaceText[],
    options?: ReplaceTextOptions,
  ) => (element: XMLDocument | Element): void => {
    const replaceTexts = GeneralHelper.arrayify(replaceText);

    new TextReplaceHelper(options, element as XMLDocument)
      .isolateTaggedNodes()
      .applyReplacements(replaceTexts);
  };

  /**
   * Set position and size of modified shape.
   */
  static setPosition = (pos: ShapeCoordinates) => (
    element: XMLDocument | Element,
  ): void => {
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

    const xfrm = element.getElementsByTagName('a:off')[0].parentNode as Element;

    Object.keys(pos).forEach((key) => {
      const value = Math.round(pos[key]);
      if(typeof value !== 'number') return;

      xfrm
        .getElementsByTagName(map[key].tag)[0]
        .setAttribute(map[key].attribute, value);
    });
  };
}
