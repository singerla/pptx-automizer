import { XmlHelper } from './xml-helper';
import { XmlDocument, XmlElement } from '../types/xml-types';

export default class ModifyHelper {
  /**
   * Set value of an attribute.
   * @param tagName specify the tag name to search for
   * @param attribute name of target attribute
   * @param value the value to be set on the attribute
   * @param [count] specify if element index is different to zero
   */
  static setAttribute =
    (
      tagName: string,
      attribute: string,
      value: string | number,
      count?: number,
    ) =>
    (element: XmlDocument): void => {
      const item = element.getElementsByTagName(tagName)[count || 0];
      if (item.setAttribute !== undefined) {
        item.setAttribute(attribute, String(value));
      }
    };

  /**
   * Dump current element to console.
   */
  static dump = (element: XmlDocument | XmlElement): void => {
    XmlHelper.dump(element);
  };

  /**
   * Dump current chart to console.
   */
  static dumpChart = (
    element: XmlDocument | XmlElement,
    chart: XmlDocument,
  ): void => {
    XmlHelper.dump(chart);
  };
}
