import {
  ElementInfo,
  ElementType,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { vd } from './general-helper';

export const nsMain =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
export const mapUriType = {
  'http://schemas.openxmlformats.org/drawingml/2006/table': 'table',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'chart',
};

/**
 * Class that represents an XML slide helper
 */
export class XmlSlideHelper {
  private slideXml: XmlDocument;

  /**
   * Constructor for the XmlSlideHelper class.
   * @param {XmlDocument} slideXml - The slide XML document to be used by the helper.
   */
  constructor(slideXml: XmlDocument) {
    if (!slideXml) {
      throw Error('Slide XML is not defined');
    }
    this.slideXml = slideXml;
  }

  getSlideCreationId(): number | undefined {
    const creationIdItem = this.slideXml
      .getElementsByTagName('p14:creationId')
      .item(0);

    if (!creationIdItem) {
      return;
    }

    const creationIdSlide = creationIdItem.getAttribute('val');
    if (!creationIdSlide) {
      return;
    }

    return Number(creationIdSlide);
  }

  /**
   * Get an array of ElementInfo objects for all named elements on a slide.
   * @param filterTags Use an array of strings to filter the output array
   */
  getAllElements(filterTags?: string[]): ElementInfo[] {
    const elementInfo: ElementInfo[] = [];
    try {
      const shapeNodes = this.getNamedElements(filterTags);
      shapeNodes.forEach((shapeNode: XmlElement) => {
        elementInfo.push(XmlSlideHelper.getElementInfo(shapeNode));
      });
    } catch (error) {
      console.error(error);
      throw new Error(`Failed to retrieve elements: ${error.message}`);
    }

    return elementInfo;
  }

  /**
   * Get all text element IDs from the slide.
   * @return {string[]} An array of text element IDs.
   */
  getAllTextElementIds(useCreationIds?: boolean): string[] {
    const elementIds: string[] = [];

    try {
      elementIds.push(
        ...this.getAllElements(['sp'])
          .filter((element) => element.hasTextBody)
          .map((element) => (useCreationIds ? element.id : element.name)),
      );
    } catch (error) {
      console.error(error);
      throw new Error(`Failed to retrieve text element IDs: ${error.message}`);
    }

    return elementIds;
  }

  static getElementInfo(slideElement: XmlElement): ElementInfo {
    return {
      name: XmlSlideHelper.getElementName(slideElement),
      id: XmlSlideHelper.getElementCreationId(slideElement),
      type: XmlSlideHelper.getElementType(slideElement),
      position: XmlSlideHelper.parseShapeCoordinates(slideElement),
      hasTextBody: !!XmlSlideHelper.getTextBody(slideElement),
      getXmlElement: () => slideElement,
    };
  }

  /**
   * Retreives a list of all named elements on a slide. Automation requires at least a name.
   * @param filterTags Use an array of strings to filter the output array
   */
  getNamedElements(filterTags?: string[]): XmlElement[] {
    const skipTags = ['spTree'];

    const nvPrs = this.slideXml.getElementsByTagNameNS(nsMain, 'cNvPr');
    const namedElements = <XmlElement[]>[];
    XmlHelper.modifyCollection(nvPrs, (nvPr: any) => {
      const parentNode = nvPr.parentNode.parentNode;
      const parentTag = parentNode.localName;
      if (
        !skipTags.includes(parentTag) &&
        (!filterTags?.length || filterTags.includes(parentTag))
      ) {
        namedElements.push(parentNode);
      }
    });
    return namedElements;
  }

  static getTextBody(shapeNode: XmlElement): XmlElement {
    return shapeNode.getElementsByTagNameNS(nsMain, 'txBody').item(0);
  }

  static getNonVisibleProperties(shapeNode: XmlElement): XmlElement {
    return shapeNode.getElementsByTagNameNS(nsMain, 'cNvPr').item(0);
  }

  static getElementName(slideElement: XmlElement) {
    const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
    if (cNvPr) {
      return cNvPr.getAttribute('name');
    }
  }

  static getElementCreationId(slideElement: XmlElement): string | undefined {
    const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
    if (cNvPr) {
      const creationIdElement = cNvPr
        .getElementsByTagName('a16:creationId')
        .item(0);

      if (creationIdElement) {
        return creationIdElement.getAttribute('id');
      }
    }
  }

  /**
   * Parses local tag name to specify element type in case it is a 'graphicFrame'.
   * @param slideElementParent
   */
  static getElementType(slideElementParent: XmlElement): ElementType {
    let type = slideElementParent.localName;

    switch (type) {
      case 'graphicFrame':
        const graphicData =
          slideElementParent.getElementsByTagName('a:graphicData')[0];
        const uri = graphicData.getAttribute('uri');
        type = mapUriType[uri] ? mapUriType[uri] : type;
        break;
    }

    return type as ElementType;
  }

  static parseShapeCoordinates(slideElementParent: XmlElement) {
    const xFrmsA = slideElementParent.getElementsByTagName('a:xfrm');
    const xFrmsP = slideElementParent.getElementsByTagName('p:xfrm');
    const xFrms = xFrmsP.item(0) ? xFrmsP : xFrmsA;

    const position = {
      x: 0,
      y: 0,
      cx: 0,
      cy: 0,
    };

    if (!xFrms.item(0)) {
      return position;
    }

    const xFrm = xFrms.item(0);
    const Off = xFrm.getElementsByTagName('a:off').item(0);
    const Ext = xFrm.getElementsByTagName('a:ext').item(0);

    position.x = XmlSlideHelper.parseCoordinate(Off, 'x');
    position.y = XmlSlideHelper.parseCoordinate(Off, 'y');
    position.cx = XmlSlideHelper.parseCoordinate(Ext, 'cx');
    position.cy = XmlSlideHelper.parseCoordinate(Ext, 'cy');

    return position;
  }

  static parseCoordinate = (
    element: XmlElement,
    attributeName: string,
  ): number => {
    return Number(element.getAttribute(attributeName));
  };
}
