import { Color } from '../types/modify-types';
import { XmlHelper } from './xml-helper';

export type XmlElementParams = {
  color?: Color;
};

export default class XmlElements {
  element: XMLDocument | Element;
  document: XMLDocument;
  params: XmlElementParams;

  constructor(element: XMLDocument | Element, params?: XmlElementParams) {
    this.element = element;
    this.document = element.ownerDocument;
    this.params = params;
  }

  text(): this {
    const r = this.document.createElement('a:r');
    r.appendChild(this.textRangeProps());
    r.appendChild(this.textContent());

    const previousSibling = this.element.getElementsByTagName('a:pPr')[0];
    XmlHelper.insertAfter(r, previousSibling);

    return this;
  }

  textRangeProps() {
    const rPr = this.document.createElement('a:rPr');
    const endParaRPr = this.element.getElementsByTagName('a:endParaRPr')[0];
    rPr.setAttribute('lang', endParaRPr.getAttribute('lang'));
    rPr.setAttribute('sz', endParaRPr.getAttribute('sz'));

    rPr.appendChild(this.line());
    rPr.appendChild(this.effectLst());
    rPr.appendChild(this.lineTexture());
    rPr.appendChild(this.fillTexture());

    return rPr;
  }

  textContent(): Element {
    const t = this.document.createElement('a:t');
    t.textContent = ' ';
    return t;
  }

  effectLst(): Element {
    return this.document.createElement('a:effectLst');
  }

  lineTexture(): Element {
    return this.document.createElement('a:uLnTx');
  }

  fillTexture(): Element {
    return this.document.createElement('a:uFillTx');
  }

  line(): Element {
    const ln = this.document.createElement('a:ln');
    const noFill = this.document.createElement('a:noFill');
    ln.appendChild(noFill);
    return ln;
  }

  solidFill(): this {
    const solidFill = this.document.createElement('a:solidFill');
    const colorType = this.colorType();
    solidFill.appendChild(colorType);

    this.element.appendChild(solidFill);

    return this;
  }

  colorType() {
    const tag = 'a:' + this.params.color.type;
    const colorType = this.document.createElement(tag);
    this.colorValue(colorType);
    return colorType;
  }

  colorValue(colorType: Element) {
    colorType.setAttribute('val', this.params.color.value);
  }
}
