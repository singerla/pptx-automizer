import { XmlDocument } from '../types/xml-types';
import ModifyColorHelper from './modify-color-helper';
import { Color } from '../types/modify-types';
import { XmlHelper } from './xml-helper';
import { vd } from './general-helper';
import { IMaster, ModifyImageHelper } from '../index';

export default class ModifyBackgroundHelper {
  /**
   * Set solid fill of master background
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

  /**
   * Modify a slideMaster background image's relation target
   * @param master
   * @param imageName
   */
  static setRelationTarget = (master: IMaster, imageName: string) => {
    let targetRelation = '';
    master.modify((masterXml) => {
      targetRelation =
        ModifyBackgroundHelper.getBackgroundProperties(masterXml);
    });
    master.modifyRelations((relXml) => {
      const relations = XmlHelper.findByAttributeValue(
        relXml.getElementsByTagName('Relationship'),
        'Id',
        targetRelation,
      );
      if (relations[0]) {
        ModifyImageHelper.setRelationTarget(imageName)(undefined, relations[0]);
      }
    });
  };

  /**
   * Extract background properties from slideMaster xml
   */
  static getBackgroundProperties = (slideMasterXml: XmlDocument): string => {
    const bgPr = slideMasterXml.getElementsByTagName('p:bgPr')?.item(0);
    if (bgPr) {
      const blip = bgPr
        .getElementsByTagName('a:blip')
        ?.item(0)
        .getAttribute('r:embed');
      return blip;
    } else {
      throw 'No background properties for slideMaster';
    }
  };
}
