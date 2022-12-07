import { XmlHelper } from './xml-helper';
import { vd } from './general-helper';

export default class ModifyPresentationHelper {
  /**
   * Get Collection of slides
   */
  static getSlidesCollection = (xml: XMLDocument) => {
    return xml.getElementsByTagName('p:sldId');
  };

  /**
   * Pass an array of slide numbers to define a target sort order.
   * First slide starts by 1.
   * @order Array of slide numbers, starting by 1
   */
  static sortSlides = (order: number[]) => (xml: XMLDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    order.map((index, i) => order[i]--);
    XmlHelper.sortCollection(slides, order);
  };

  /**
   * Set ids to prevent corrupted pptx.
   * Must start with 256 and increment by one.
   */
  static normalizeSlideIds = (xml: XMLDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    const firstId = 256;
    XmlHelper.modifyCollection(slides, (slide: Element, i) => {
      slide.setAttribute('id', String(firstId + i));
    });
  };
}
