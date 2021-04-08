import { XmlHelper } from './xml-helper';

export default class ModifyHelper {
  /**
   * Set value of an attribute.
   * @param tagName specify the tag name to search for
   * @param attribute name of target attribute
   * @param value the value to be set on the attribute
   * @param [count] specify if element index is different to zero
   */
  static setAttribute = (
    tagName: string,
    attribute: string,
    value: string | number,
    count?: number,
  ) => (element: XMLDocument): void => {
    element
      .getElementsByTagName(tagName)
      [count || 0].setAttribute(attribute, String(value));
  };

  /**
   * Dump current element to console.
   */
  static dump = (element: XMLDocument | Document | Element): void => {
    XmlHelper.dump(element);
  };
}
