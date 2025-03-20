import { XmlElement } from '../types/xml-types';
import { ImageStyle } from '../types/modify-types';
import slugify from 'slugify';

export default class ModifyImageHelper {
  /**
   * Update the "Target" attribute of a created image relation.
   * This will change the image itself. Load images with Automizer.loadMedia
   * @param filename name of target image in root template media folder.
   */
  static setRelationTarget = (filename: string) => {
    return (element: XmlElement, arg1: XmlElement): void => {
      arg1.setAttribute('Target', '../media/' + slugify(filename));
    };
  };

  /*
    Update an existing duotone image overlay element (WIP)
    Apply a duotone color to an image p:blipFill -> a:blip fill element.
    Works best on white icons, see __tests__/media/feather.png
   */
  static setDuotoneFill =
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
