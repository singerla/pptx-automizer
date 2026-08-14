import { BulletListContent, Color, ShapeOutline } from '../types/modify-types';
import { XmlHelper } from './xml-helper';
import { DOMParser } from '@xmldom/xmldom';
import type { Node } from '@xmldom/xmldom';
import { lnLRTB } from './xml/lnLRTB';
import { XmlDocument, XmlElement } from '../types/xml-types';

export type XmlElementParams = {
  color?: Color;
  outline?: ShapeOutline;
};

/**
 * Children of <a:ln>, in OOXML schema order (CT_LineProperties).
 * Required to insert new children at a valid position: a wrong sequence
 * makes PowerPoint ask to repair the file.
 */
const LINE_FILL_TAGS = [
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:pattFill',
];
const LINE_DASH_TAGS = ['a:prstDash', 'a:custDash'];
const LINE_AFTER_DASH_TAGS = [
  'a:round',
  'a:bevel',
  'a:miter',
  'a:headEnd',
  'a:tailEnd',
  'a:extLst',
];

/**
 * Children of <c:ser>, in OOXML schema order — the union of all CT_*Ser
 * variants (bar, line, pie, scatter, radar, area, bubble, surface); their
 * shared members never disagree on relative order.
 */
export const C_SER_CHILD_ORDER = [
  'c:idx',
  'c:order',
  'c:tx',
  'c:spPr',
  'c:invertIfNegative',
  'c:pictureOptions',
  'c:explosion',
  'c:marker',
  'c:dPt',
  'c:dLbls',
  'c:trendline',
  'c:errBars',
  'c:cat',
  'c:val',
  'c:xVal',
  'c:yVal',
  'c:smooth',
  'c:shape',
  'c:bubbleSize',
  'c:bubble3D',
  'c:extLst',
];

/**
 * Children of <c:dPt> (CT_DPt), in OOXML schema order. Note that it differs
 * from CT_Ser: inside a data point, `c:bubble3D` comes *before* `c:spPr`.
 */
export const C_DPT_CHILD_ORDER = [
  'c:idx',
  'c:invertIfNegative',
  'c:marker',
  'c:bubble3D',
  'c:explosion',
  'c:spPr',
  'c:pictureOptions',
  'c:extLst',
];

/**
 * Children of <c:spPr>/<a:spPr> (CT_ShapeProperties), in OOXML schema order.
 */
export const A_SPPR_CHILD_ORDER = [
  'a:xfrm',
  'a:custGeom',
  'a:prstGeom',
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
  'a:ln',
  'a:effectLst',
  'a:effectDag',
  'a:scene3d',
  'a:sp3d',
  'a:extLst',
];

/**
 * Children of <c:dLbls> (CT_DLbls), in OOXML schema order: the sparse
 * <c:dLbl> point overrides come first, followed by the group-level settings.
 */
export const C_DLBLS_CHILD_ORDER = [
  'c:dLbl',
  'c:delete',
  'c:numFmt',
  'c:spPr',
  'c:txPr',
  'c:dLblPos',
  'c:showLegendKey',
  'c:showVal',
  'c:showCatName',
  'c:showSerName',
  'c:showPercent',
  'c:showBubbleSize',
  'c:separator',
  'c:showLeaderLines',
  'c:leaderLines',
  'c:extLst',
];

export default class XmlElements {
  element: XmlDocument | XmlElement;

  document: XmlDocument;
  params: XmlElementParams;
  defaultValues: Record<string, string>;
  paragraphTemplate: XmlElement;
  runTemplate: XmlElement;

  constructor(element: XmlDocument | XmlElement, params?: XmlElementParams) {
    this.element = element;
    this.document = element.ownerDocument;
    this.params = params;
    this.defaultValues = {
      color: 'CCCCCC',
      size: '1000',
    };
  }

  text(): this {
    const r = this.document.createElement('a:r');
    r.appendChild(this.textRangeProps());
    r.appendChild(this.textContent());

    let paragraphProps = this.element.getElementsByTagName('a:pPr').item(0);

    if (!paragraphProps) {
      paragraphProps = this.paragraphProps();
    }

    XmlHelper.insertAfter(r, paragraphProps);

    return this;
  }

  createTextBody(): XmlElement {
    let txBody = this.element.getElementsByTagName('p:txBody')[0];
    if (!txBody) {
      txBody = this.document.createElement('p:txBody');
      this.element.appendChild(txBody);

      const bodyPr = this.document.createElement('a:bodyPr');
      txBody.appendChild(bodyPr);

      const lstStyle = this.document.createElement('a:lstStyle');
      txBody.appendChild(lstStyle);

      this.paragraphTemplate = this.document.createElement('a:p');
      txBody.appendChild(this.paragraphTemplate);

      this.runTemplate = this.document.createElement('a:r');
      const rPr = this.document.createElement('a:rPr');
      this.runTemplate.appendChild(rPr);
    } else {
      let bodyPr = txBody.getElementsByTagName('a:bodyPr')[0];
      if (!bodyPr) {
        bodyPr = this.document.createElement('a:bodyPr');
        txBody.insertBefore(bodyPr, txBody.firstChild);
      }

      let lstStyle = txBody.getElementsByTagName('a:lstStyle')[0];
      if (!lstStyle) {
        lstStyle = this.document.createElement('a:lstStyle');
        txBody.insertBefore(lstStyle, bodyPr.nextSibling);
      }

      const paragraphs = txBody.getElementsByTagName('a:p');
      this.paragraphTemplate = paragraphs[0];
      XmlHelper.sliceCollection(paragraphs, 0);

      const runs = this.paragraphTemplate.getElementsByTagName('a:r');
      if (runs.length > 0) {
        this.runTemplate = runs[0];
      } else {
        this.runTemplate = this.document.createElement('a:r');
        const rPr = this.document.createElement('a:rPr');
        this.runTemplate.appendChild(rPr);
      }
    }
    return txBody;
  }

  createBodyProperties(txBody: XmlElement): XmlElement {
    const bodyPr = this.document.createElement('a:bodyPr');
    txBody.appendChild(bodyPr);
    return bodyPr;
  }

  addBulletList(list: BulletListContent): void {
    const txBody = this.createTextBody();
    this.createBodyProperties(txBody);
    this.processList(txBody, list, 0);
  }

  processList(txBody: XmlElement, items: BulletListContent, level: number): void {
    items.forEach((item) => {
      if (Array.isArray(item)) {
        this.processList(txBody, item, level + 1);
      } else {
        const p = this.createParagraph(level);
        const r = this.createTextRun(String(item));
        p.appendChild(r);
        txBody.appendChild(p);
      }
    });
  }

  createParagraph(level: number): XmlElement {
    const p = this.paragraphTemplate.cloneNode(true) as XmlElement;
    const pPr = p.getElementsByTagName('a:pPr')[0];
    if (pPr) {
      if (level > 0) {
        pPr.setAttribute('lvl', String(level));
        pPr.removeAttribute('indent');
        pPr.removeAttribute('marL');
      } else {
        pPr.removeAttribute('lvl');
      }
    } else {
      const newPPr = this.document.createElement('a:pPr');
      if (level > 0) {
        newPPr.setAttribute('lvl', String(level));
      }
      p.insertBefore(newPPr, p.firstChild);
    }
    const runs = p.getElementsByTagName('a:r');
    XmlHelper.sliceCollection(runs, 0);
    return p;
  }

  createTextRun(text: string): XmlElement {
    const r = this.runTemplate.cloneNode(true) as XmlElement;
    const t = r.getElementsByTagName('a:t')[0];
    if (t) {
      t.textContent = XmlHelper.sanitizeText(text);
    } else {
      const newT = this.document.createElement('a:t');
      newT.textContent = XmlHelper.sanitizeText(text);
      r.appendChild(newT);
    }

    return r;
  }

  paragraphProps() {
    const p = this.element.getElementsByTagName('a:p').item(0);
    p.appendChild(this.document.createElement('a:pPr'));
    const paragraphRangeProps = this.element
      .getElementsByTagName('a:pPr')
      .item(0);

    const endParaRPr = this.element
      .getElementsByTagName('a:endParaRPr')
      .item(0);
    XmlHelper.moveChild(endParaRPr);

    return paragraphRangeProps;
  }

  textRangeProps() {
    const rPr = this.document.createElement('a:rPr');
    const endParaRPr = this.element.getElementsByTagName('a:endParaRPr')[0];
    rPr.setAttribute('lang', endParaRPr.getAttribute('lang'));
    rPr.setAttribute(
      'sz',
      endParaRPr.getAttribute('sz') || this.defaultValues.size,
    );

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

  /**
   * Create an <a:ln> shape outline from this.params.outline
   */
  outline(): XmlElement {
    const ln = this.document.createElement('a:ln');
    return this.applyOutline(ln);
  }

  /**
   * Apply this.params.outline to an existing (or freshly created) <a:ln>.
   * Only given properties are touched, the rest is left to the template.
   * Insertion respects the schema sequence of CT_LineProperties:
   * fill -> dash -> join -> head/tailEnd -> extLst
   *
   * @param ln - The <a:ln> element to update
   */
  applyOutline(ln: XmlElement): XmlElement {
    const outline = this.params?.outline;
    if (!outline) return ln;

    if (outline.weight !== undefined) {
      ln.setAttribute('w', String(Math.round(outline.weight)));
    }

    if (outline.color) {
      this.params.color = outline.color;
      const solidFill = this.solidFill();
      const currentFill = XmlHelper.getFirstDirectChild(ln, LINE_FILL_TAGS);

      if (currentFill) {
        ln.replaceChild(solidFill, currentFill);
      } else {
        ln.insertBefore(solidFill, ln.firstChild);
      }
    }

    if (outline.type) {
      const prstDash = this.prstDash();
      prstDash.setAttribute('val', outline.type);
      const currentDash = XmlHelper.getFirstDirectChild(ln, LINE_DASH_TAGS);

      if (currentDash) {
        ln.replaceChild(prstDash, currentDash);
      } else {
        const anchor = XmlHelper.getFirstDirectChild(ln, LINE_AFTER_DASH_TAGS);
        if (anchor) {
          ln.insertBefore(prstDash, anchor);
        } else {
          ln.appendChild(prstDash);
        }
      }
    }

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
    colorType.setAttribute(
      'val',
      this.params?.color?.value || this.defaultValues.color,
    );

    if (this.params?.color?.alpha !== undefined) {
      const alpha = this.document.createElement('a:alpha');
      const rawAlpha = Number(this.params.color.alpha);
      // Normalize alpha to OOXML thousandths of percent (0-100000):
      // 0-1 (exclusive): fraction (e.g. 0.5 → 50000)
      // 1-100: percentage (e.g. 50 → 50000)
      // >100: already in thousandths of percent
      let alphaVal: number;
      if (rawAlpha > 0 && rawAlpha < 1) {
        alphaVal = Math.round(rawAlpha * 100000);
      } else if (rawAlpha >= 1 && rawAlpha <= 100) {
        alphaVal = Math.round(rawAlpha * 1000);
      } else {
        alphaVal = Math.round(rawAlpha);
      }
      alpha.setAttribute('val', String(alphaVal));
      colorType.appendChild(alpha);
    }
  }

  /**
   * A minimal <c:dPt> shell: `c:idx`, `c:invertIfNegative` and `c:bubble3D`
   * (which PowerPoint always writes) and nothing else. A data point carries
   * no formatting the caller did not ask for — modifications create `c:spPr`
   * etc. on demand. `c:invertIfNegative` must be written explicitly: OOXML
   * defaults the absent element to *true*, which makes PowerPoint invert the
   * fill of negative-value bars (white with a border) — silently overriding
   * the very fill the caller styled the point for. LibreOffice ignores the
   * flag, so pixel-based golden decks cannot catch this.
   */
  buildDataPoint(): XmlElement {
    const dPt = this.document.createElement('c:dPt');
    dPt.appendChild(this.idx());
    const invertIfNegative = this.document.createElement('c:invertIfNegative');
    invertIfNegative.setAttribute('val', '0');
    dPt.appendChild(invertIfNegative);
    const bubble3D = this.document.createElement('c:bubble3D');
    bubble3D.setAttribute('val', '0');
    dPt.appendChild(bubble3D);
    return dPt;
  }

  dataPoint(): this {
    XmlHelper.insertInSchemaOrder(
      this.element as XmlElement,
      this.buildDataPoint(),
      C_SER_CHILD_ORDER,
    );
    return this;
  }

  /**
   * An empty <c:spPr> shell in schema position: all children of
   * CT_ShapeProperties are optional, and an absent property inherits from the
   * series/theme defaults. Fills, lines etc. are added by the modifications
   * that asked for them — the former grey solidFill + <a:ln><a:noFill/>
   * default erased the segments of line charts.
   */
  shapeProperties() {
    const spPr = this.document.createElement('c:spPr');
    const parentTag = (this.element as XmlElement).nodeName;
    const order =
      parentTag === 'c:dPt' ? C_DPT_CHILD_ORDER : C_SER_CHILD_ORDER;
    XmlHelper.insertInSchemaOrder(this.element as XmlElement, spPr, order);
  }

  /**
   * A bare <a:ln> in schema position — no fabricated <a:noFill>. Border
   * modifications fill it with the properties the caller asked for.
   */
  plainLine() {
    const ln = this.document.createElement('a:ln');
    XmlHelper.insertInSchemaOrder(
      this.element as XmlElement,
      ln,
      A_SPPR_CHILD_ORDER,
    );
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

  /**
   * An empty <c:dLbls>: all children of CT_DLbls are optional, and every
   * label property (visibility included) is inherited from the chart's
   * defaults. No fabricated formatting, no forced `showVal`.
   */
  dataPointLabels() {
    const dLbls = this.document.createElement('c:dLbls');
    XmlHelper.insertInSchemaOrder(
      this.element as XmlElement,
      dLbls,
      C_SER_CHILD_ORDER,
    );
  }

  /**
   * A minimal <c:dLbl> point override: `c:idx` plus an empty `c:txPr`
   * scaffold so that text styling modifications have a target — without
   * opinionated defaults (no size, no fill, no `showVal`).
   */
  buildDataPointLabel(): XmlElement {
    const dLbl = this.document.createElement('c:dLbl');
    dLbl.appendChild(this.idx());

    const txPr = this.document.createElement('c:txPr');
    txPr.appendChild(this.document.createElement('a:bodyPr'));
    txPr.appendChild(this.document.createElement('a:lstStyle'));

    const p = this.document.createElement('a:p');
    const pPr = this.document.createElement('a:pPr');
    pPr.appendChild(this.document.createElement('a:defRPr'));
    p.appendChild(pPr);
    const endParaRPr = this.document.createElement('a:endParaRPr');
    endParaRPr.setAttribute('lang', 'en-US');
    p.appendChild(endParaRPr);
    txPr.appendChild(p);

    dLbl.appendChild(txPr);
    return dLbl;
  }

  dataPointLabel() {
    XmlHelper.insertInSchemaOrder(
      this.element as XmlElement,
      this.buildDataPointLabel(),
      C_DLBLS_CHILD_ORDER,
    );
  }

  tableCellBorder(tag: 'a:lnL' | 'a:lnR' | 'a:lnT' | 'a:lnB') {
    const doc = new DOMParser().parseFromString(lnLRTB, 'application/xml');
    const ele = doc.getElementsByTagName(tag)[0] as unknown as Node;
    const firstChild = this.element.firstChild;
    this.element.insertBefore(ele.cloneNode(true), firstChild);
  }
}
