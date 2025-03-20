import {
  ElementInfo,
  ElementType,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import HasShapes from '../classes/has-shapes';
import { FindElementSelector, ShapeModificationCallback } from '../types/types';
import ModifyTableHelper from './modify-table-helper';
import { TableData, TableInfo } from '../types/table-types';

export const nsMain =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
export const mapUriType = {
  'http://schemas.openxmlformats.org/drawingml/2006/table': 'table',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'chart',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'chartEx',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject': 'oleObject',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink': 'hyperlink',
};

/**
 * Class that represents an XML slide helper
 */
export class XmlSlideHelper {
  private slideXml: XmlDocument;
  protected hasShapes: HasShapes;

  /**
   * Constructor for the XmlSlideHelper class.
   * @param {XmlDocument} slideXml - The slide XML document to be used by the helper.
   * @param hasShapes
   */
  constructor(slideXml: XmlDocument, hasShapes?: HasShapes) {
    if (!slideXml) {
      throw Error('Slide XML is not defined');
    }
    this.slideXml = slideXml;
    this.hasShapes = hasShapes;
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
   * @param selector
   */
  async getElement(selector: string): Promise<ElementInfo> {
    const shapeNode = XmlHelper.isElementCreationId(selector)
      ? XmlHelper.findByCreationId(this.slideXml, selector)
      : XmlHelper.findByName(this.slideXml, selector);

    return XmlSlideHelper.getElementInfo(shapeNode);
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
      getText: () => XmlSlideHelper.parseTextFragments(slideElement),
      getTableInfo: () => XmlSlideHelper.readTableInfo(slideElement),
      getAltText: () => XmlSlideHelper.getImageAltText(slideElement),
    };
  }

  /**
   * Retrieves a list of all named elements on a slide. Automation requires at least a name.
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

  static parseTextFragments(shapeNode: XmlElement): string[] {
    const txBody = XmlSlideHelper.getTextBody(shapeNode);
    const textFragments: string[] = [];
    const texts = txBody.getElementsByTagName('a:t');
    for (let t = 0; t < texts.length; t++) {
      textFragments.push(texts.item(t).textContent);
    }
    return textFragments;
  }

  static getNonVisibleProperties(shapeNode: XmlElement): XmlElement {
    return shapeNode.getElementsByTagNameNS(nsMain, 'cNvPr').item(0);
  }

  static getImageAltText(slideElement: XmlElement) {
    const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
    if (cNvPr) {
      return cNvPr.getAttribute('descr');
    }
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
      case 'oleObj':
        type = 'OLEObject';
        break;
    }

    // Check for hyperlinks
    const hasHyperlink = slideElementParent.getElementsByTagName('a:hlinkClick');
    if (hasHyperlink.length > 0) {
      type = 'Hyperlink';
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
    return parseInt(element.getAttribute(attributeName), 10);
  };

  /**
   * Asynchronously retrieves the dimensions of a slide.
   * Tries to find the dimensions from the slide XML, then from the layout, master, and presentation XMLs in order.
   *
   * @returns {Promise<{ width: number, height: number }>} The dimensions of the slide.
   * @throws Error if unable to determine dimensions.
   */
  async getDimensions(): Promise<{ width: number; height: number }> {
    try {
      const dimensions = await this.getAndExtractDimensions(
        'ppt/presentation.xml',
      );
      if (dimensions) return dimensions;
    } catch (error) {
      console.error(`Error while fetching slide dimensions: ${error}`);
      throw error;
    }
  }

  /**
   * Fetches an XML file from the given path and extracts the dimensions.
   *
   * @param {string} path - The path of the XML file.
   * @returns {Promise<{ width: number; height: number } | null>} - A promise that resolves with an object containing the width and height, or `null` if there was an error.
   */
  getAndExtractDimensions = async (
    path: string,
  ): Promise<{ width: number; height: number } | null> => {
    try {
      const xml = await XmlHelper.getXmlFromArchive(
        this.hasShapes.sourceTemplate.archive,
        path,
      );
      if (!xml) return null;

      const sldSz = xml.getElementsByTagName('p:sldSz')[0];
      if (sldSz) {
        const width = XmlSlideHelper.parseCoordinate(sldSz, 'cx');
        const height = XmlSlideHelper.parseCoordinate(sldSz, 'cy');
        return { width, height };
      }
      return null;
    } catch (error) {
      console.warn(`Error while fetching XML from path ${path}: ${error}`);
      return null;
    }
  };

  static readTableInfo = (element: XmlElement): TableInfo[] => {
    const info = <TableInfo[]>[];
    const rows = element.getElementsByTagName('a:tr');
    if (!rows) {
      console.error("Can't find a table row.");
      return info;
    }

    for (let r = 0; r < rows.length; r++) {
      const row = rows.item(r);
      const columns = row.getElementsByTagName('a:tc');
      for (let c = 0; c < columns.length; c++) {
        const cell = columns.item(c);
        const gridSpan = cell.getAttribute('gridSpan');
        const hMerge = cell.getAttribute('hMerge');
        const texts = cell.getElementsByTagName('a:t');
        const text: string[] = [];
        for (let t = 0; t < texts.length; t++) {
          text.push(texts.item(t).textContent);
        }
        info.push({
          row: r,
          column: c,
          rowXml: row,
          columnXml: cell,
          text: text,
          textContent: text.join(''),
          gridSpan: Number(gridSpan),
          hMerge: Number(hMerge),
        });
      }
    }
    return info;
  };
}
