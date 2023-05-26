import { Border, Color } from '../types/modify-types';
import { XmlHelper } from './xml-helper';
import { DOMParser } from '@xmldom/xmldom';
import { dLblXml } from './xml/dLbl';
import { lnLRTB } from './xml/lnLRTB';
import { XmlDocument, XmlElement } from '../types/xml-types';

export type XmlElementParams = {
  color?: Color;
};

export default class XmlElements {
  element: XmlDocument | XmlElement;

  document: XmlDocument;
  params: XmlElementParams;

  constructor(element: XmlDocument | XmlElement, params?: XmlElementParams) {
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

  textContent(): XmlElement {
    const t = this.document.createElement('a:t');
    t.textContent = ' ';
    return t;
  }

  effectLst(): XmlElement {
    return this.document.createElement('a:effectLst');
  }

  lineTexture(): XmlElement {
    return this.document.createElement('a:uLnTx');
  }

  fillTexture(): XmlElement {
    return this.document.createElement('a:uFillTx');
  }

  line(): XmlElement {
    const ln = this.document.createElement('a:ln');
    const noFill = this.document.createElement('a:noFill');
    ln.appendChild(noFill);
    return ln;
  }

  solidFill(): XmlElement {
    const solidFill = this.document.createElement('a:solidFill');
    const colorType = this.colorType();
    solidFill.appendChild(colorType);
    return solidFill;
  }

  colorType(): XmlElement {
    const tag = 'a:' + (this.params?.color?.type || 'srgbClr');
    const colorType = this.document.createElement(tag);
    this.colorValue(colorType);
    return colorType;
  }

  colorValue(colorType: XmlElement) {
    colorType.setAttribute('val', this.params?.color?.value || 'cccccc');
  }

  dataPoint(): this {
    const dPt = this.document.createElement('c:dPt');
    dPt.appendChild(this.idx());
    dPt.appendChild(this.spPr());

    const nextSibling = this.element.getElementsByTagName('c:cat')[0];
    if (nextSibling) {
      nextSibling.parentNode.insertBefore(dPt, nextSibling);
    }

    return this;
  }

  spPr(): XmlElement {
    const spPr = this.document.createElement('c:spPr');
    spPr.appendChild(this.solidFill());
    spPr.appendChild(this.line());
    spPr.appendChild(this.effectLst());

    return spPr;
  }

  idx(): XmlElement {
    const idx = this.document.createElement('c:idx');
    idx.setAttribute('val', String(0));
    return idx;
  }

  cellBorder(tag: 'lnL' | 'lnR' | 'lnT' | 'lnB'): this {
    const border = this.document.createElement(tag);

    border.appendChild(this.solidFill());
    border.appendChild(this.prstDash());
    border.appendChild(this.round());
    border.appendChild(this.lineEnd('headEnd'));
    border.appendChild(this.lineEnd('tailEnd'));

    return this;
  }

  prstDash() {
    const prstDash = this.document.createElement('a:prstDash');
    prstDash.setAttribute('val', 'solid');
    return prstDash;
  }

  round() {
    const round = this.document.createElement('a:round');
    return round;
  }

  lineEnd(type: 'headEnd' | 'tailEnd') {
    const lineEnd = this.document.createElement(type);
    lineEnd.setAttribute('type', 'none');
    lineEnd.setAttribute('w', 'med');
    lineEnd.setAttribute('len', 'med');
    return lineEnd;
  }

  shapeProperties() {
    const spPr = this.spPr();
    this.element.appendChild(spPr);
  }

  dataPointLabel() {
    const doc = new DOMParser().parseFromString(dLblXml);
    const ele = doc.getElementsByTagName('c:dLbl')[0];
    const firstChild = this.element.firstChild;
    this.element.insertBefore(ele.cloneNode(true), firstChild);
  }
  tableCellBorder(tag: 'a:lnL' | 'a:lnR' | 'a:lnT' | 'a:lnB') {
    const doc = new DOMParser().parseFromString(lnLRTB);
    const ele = doc.getElementsByTagName(tag)[0];
    const firstChild = this.element.firstChild;
    this.element.insertBefore(ele.cloneNode(true), firstChild);
  }
}
