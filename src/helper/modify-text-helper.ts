import { Color } from '../types/modify-types';

export default class ModifyTextHelper {
  /**
   * Set color of text insinde an <a:rPr> element
   */
   static setColor = (element: Element, color:Color): void => {
    const mapTags = {
      schemeClr: 'a:schemeClr',
      srgbClr: 'a:srgbClr'
    }

    const doc = element.ownerDocument
    const solidFill = doc.createElement('a:solidFill')
    const colorType = doc.createElement(mapTags[color.type])
    colorType.setAttribute('val', color.value)

    element.appendChild(solidFill)
    solidFill.appendChild(colorType)

    const colorElement = element.getElementsByTagName('a:solidFill')
    if(colorElement.length > 1) {
      colorElement[0].parentNode.removeChild(colorElement[0])
    }
  };

  /**
   * Set size of text insinde an <a:rPr> element
   */
   static setSize = (element: Element, size:number): void => {
    element.setAttribute('sz', String(Math.round(size)))
  };
}
