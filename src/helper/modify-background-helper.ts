import { XmlDocument } from '../types/xml-types';
import ModifyColorHelper from './modify-color-helper';
import { Color } from '../types/modify-types';

export default class ModifyBackgroundHelper {
  /**
   * Set solid fill of modified shape
   */
  static setSolidFill =
    (color: Color) =>
    (slideMasterXml: XmlDocument): void => {
      const bgPr = slideMasterXml.getElementsByTagName('p:bgPr')?.item(0);
      if (bgPr) {
        ModifyColorHelper.solidFill(color)(bgPr);
      } else {
        throw 'No background properties for slideMaster';
      }
    };
}
