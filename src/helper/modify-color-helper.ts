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

  /*
    Update an existing duotone image overlay element (WIP)
    Apply a duotone color to an image p:blipFill -> a:blip fill element.
   */
  static duotone =
    (duotoneParams: ImageStyle['duotone']) =>
    (element: XmlElement): void => {
      const blipFill = element.getElementsByTagName('p:blipFill');
      if (!blipFill) {
        return;
      }
      const duotone = blipFill.item(0).getElementsByTagName('a:duotone')[0];
      if (duotone) {
        if (duotoneParams?.color) {
          const srgbClr = duotone.getElementsByTagName('a:srgbClr')[0];
          if (srgbClr) {
            // Only sRgb supported
            srgbClr.setAttribute('val', String(duotoneParams.color.value));

            if (duotoneParams?.tint !== undefined) {
              // tint needs to be 0 - 100000
              const tint = srgbClr.getElementsByTagName('a:tint')[0];
              if (tint) {
                tint.setAttribute('val', String(duotoneParams.tint));
              }
            }

            if (duotoneParams?.satMod !== undefined) {
              const satMod = srgbClr.getElementsByTagName('a:satMod')[0];
              if (satMod) {
                satMod.setAttribute('val', String(duotoneParams.satMod));
              }
            }
          }
        }
        if (duotoneParams?.prstClr) {
          const prstClr = duotone.getElementsByTagName('a:prstClr')[0];
          if (prstClr) {
            // Only tested with "black" and "white"
            prstClr.setAttribute('val', String(duotoneParams.prstClr));
          }
        }
      }
    };
}
