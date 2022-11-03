import { XmlHelper } from './xml-helper';
import { vd } from './general-helper';

export default class ModifyPresentationHelper {
  /**
   * Pass an array of slide numbers to define a target sort order.
   * First slide starts by 1.
   * @order Array of slide numbers, starting by 1
   */
  static sortSlides = (order: number[]) => (xml: XMLDocument) => {
    order.map((index, i) => order[i]--);
    const slides = xml.getElementsByTagName('p:sldId');
    const firstId = 256;
    XmlHelper.sortCollection(slides, order, (slide: Element, i) => {
      slide.setAttribute('id', String(firstId + i));
    });

    // XmlHelper.dump(
    //   xml.getElementsByTagName('p:sldId')[0].parentNode as Element,
    // );
  };
}
