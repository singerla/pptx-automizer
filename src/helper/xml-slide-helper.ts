import {
  ElementInfo,
  ElementPosition,
  ElementType,
  ElementVisualType,
  GroupInfo,
  LayoutInfo,
  PlaceholderInfo,
  SlideHelperProps,
  TextParagraph,
  TextParagraphGroup,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { TableInfo } from '../types/table-types';
import IArchive from '../interfaces/iarchive';
import { XmlTemplateHelper } from './xml-template-helper';
import XmlPlaceholderHelper from './xml-placeholder-helper';
import { vd } from './general-helper';
import { FindElementMultiSelector, FindElementSelector } from '../types/types';

export const nsMain =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
export const mapUriType = {
  'http://schemas.openxmlformats.org/drawingml/2006/table': 'table',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'chart',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'chartEx',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject':
    'oleObject',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
    'hyperlink',
};

/**
 * Class that represents an XML slide helper
 */
export class XmlSlideHelper {
  private slideXml: XmlDocument;
  protected sourceArchive: IArchive;
  protected slideNumber: number;
  protected sourceLayoutId: number;

  /**
   * Constructor for the XmlSlideHelper class.
   * @param {XmlDocument} slideXml - The slide XML document to be used by the helper.
   * @param params
   */
  constructor(slideXml: XmlDocument, params: SlideHelperProps) {
    if (!slideXml) {
      throw Error('Slide XML is not defined');
    }
    this.slideXml = slideXml;
    this.sourceArchive = params.sourceArchive;
    this.slideNumber = params.slideNumber;
    this.sourceLayoutId = params.sourceLayoutId;
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

  async getSlideLayout(): Promise<LayoutInfo> {
    const xml = await this.getSlideLayoutXml(this.sourceLayoutId);
    return XmlTemplateHelper.getLayoutInfo(xml);
  }

  /**
   * Get an ElementInfo object for the target element on the slide.
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
   * @param layoutPlaceholders
   */
  getAllElements(
    filterTags?: string[],
    layoutPlaceholders?: PlaceholderInfo[],
  ): ElementInfo[] {
    const elementInfo: ElementInfo[] = [];

    try {
      const shapeNodes = this.getNamedElements(filterTags);
      shapeNodes.forEach((shapeNode: XmlElement) => {
        elementInfo.push(
          XmlSlideHelper.getElementInfo(shapeNode, layoutPlaceholders),
        );
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

  static getElementInfo(
    slideElement: XmlElement,
    layoutPlaceholders?: PlaceholderInfo[],
  ): ElementInfo {
    const selector = this.getSelector(slideElement)
    const position = XmlSlideHelper.parseShapeCoordinates(slideElement);
    const type = XmlSlideHelper.getElementType(slideElement);

    return {
      name: selector.name,
      id: selector.creationId,
      creationId: selector.creationId,
      nameIdx: selector.nameIdx,
      type,
      position,
      placeholder: XmlPlaceholderHelper.getPlaceholderInfo(
        slideElement,
        layoutPlaceholders,
      ),
      hasTextBody: !!XmlSlideHelper.getTextBody(slideElement),
      getXmlElement: () => slideElement,
      getText: () => XmlSlideHelper.parseTextFragments(slideElement),
      getParagraphs: () => XmlSlideHelper.parseTextParagraphs(slideElement),
      getParagraphGroups: () =>
        XmlSlideHelper.parseParagraphGroups(slideElement),
      getTableInfo: () => XmlSlideHelper.readTableInfo(slideElement),
      getAltText: () => XmlSlideHelper.getImageAltText(slideElement),
      getGroupInfo: () => XmlSlideHelper.parseGroupInfo(slideElement),
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

    if (!txBody) {
      return textFragments;
    }

    const texts = txBody.getElementsByTagName('a:t');
    for (let t = 0; t < texts.length; t++) {
      const text = texts.item(t);
      textFragments.push(text.textContent);
    }
    return textFragments;
  }

  static parseParagraphGroups(shapeNode: XmlElement): TextParagraphGroup[] {
    const rawParagraphs = XmlSlideHelper.parseTextParagraphs(shapeNode);
    return XmlSlideHelper.groupSimilarParagraphs(rawParagraphs);
  }

  static parseTextParagraphs(shapeNode: XmlElement): TextParagraph[] {
    const textParagraphs: TextParagraph[] = [];

    // Find txBody element first
    const txBody =
      shapeNode.getElementsByTagName('p:txBody')[0] ||
      shapeNode.getElementsByTagName('a:txBody')[0];

    if (!txBody) return textParagraphs;

    // Get all paragraph elements
    const paragraphs = txBody.getElementsByTagName('a:p');

    for (const p of Array.from(paragraphs)) {
      const paragraph: TextParagraph = { texts: [] };

      // Check for paragraph properties (indent and bullet)
      const pPr = p.getElementsByTagName('a:pPr')[0];

      if (pPr) {
        XmlSlideHelper.setParagraphProperties(pPr, paragraph);
      }

      // Get all text runs in the paragraph
      const runs = p.getElementsByTagName('a:r');
      const texts: string[] = [];

      for (const run of Array.from(runs)) {
        XmlSlideHelper.setTextProperties(run, paragraph);

        // Get text content
        const textElements = run.getElementsByTagName('a:t');
        for (const textElement of Array.from(textElements)) {
          texts.push(textElement.textContent || '');
        }

        // Check if the next sibling after rPr is a line break
        const nextSibling = run.nextSibling;
        if (nextSibling && nextSibling.nodeName === 'a:br') {
          texts.push(`\n`);
        }
      }

      // Only add paragraphs that have text content
      if (texts.length > 0) {
        paragraph.texts = texts;
        textParagraphs.push(paragraph);
      }
    }

    return textParagraphs;
  }

  static setTextProperties(run: XmlElement, paragraph: TextParagraph) {
    const rPr = run.getElementsByTagName('a:rPr')[0];
    if (rPr) {
      const isBold = rPr.getAttribute('b') === '1';
      const isUnderlined = rPr.getAttribute('u');
      const isItalic = rPr.getAttribute('i') === '1';
      const fontSize = parseInt(rPr.getAttribute('sz') || '0') / 100; // Convert to points

      if (isBold) paragraph.isBold = true;
      if (isItalic) paragraph.isItalic = true;
      if (isUnderlined) paragraph.isUnderlined = true;
      if (fontSize) paragraph.fontSize = fontSize;
    }
  }

  static setParagraphProperties(pPr: XmlElement, paragraph: TextParagraph) {
    const marL = pPr.getAttribute('marL');
    if (marL) {
      paragraph.indent = parseInt(marL);
    }

    const buChar = pPr.getElementsByTagName('a:buChar')[0];
    if (buChar) {
      paragraph.bullet = buChar.getAttribute('char');
    }

    // Check for numbered list
    const buAutoNum = pPr.getElementsByTagName('a:buAutoNum')[0];
    if (buAutoNum) {
      paragraph.isNumbered = true;
      paragraph.numberingType = buAutoNum.getAttribute('type') || undefined;
      paragraph.startAt = buAutoNum.getAttribute('startAt') || undefined;
    }

    // Check for alignment
    const algn = pPr.getAttribute('algn');
    if (algn) {
      paragraph.align = algn as TextParagraph['align'];
    }
  }

  static groupSimilarParagraphs(
    paragraphs: TextParagraph[],
  ): TextParagraphGroup[] {
    const groups: TextParagraphGroup[] = [];
    let currentGroup: TextParagraphGroup | null = null;

    const getDefinedProperties = (paragraph: TextParagraph) => {
      const properties: Record<string, any> = {};

      const propertyKeys = [
        'fontSize',
        'isBold',
        'isItalic',
        'isUnderlined',
        // 'indent',
        'align',
        'isNumbered',
        'numberingType',
        'bullet',
        'startAt',
        'breaks',
      ] as const;

      for (const key of propertyKeys) {
        if (paragraph[key] !== undefined) {
          properties[key] = paragraph[key];
        }
      }

      return properties;
    };

    for (const paragraph of paragraphs) {
      const properties = getDefinedProperties(paragraph);

      // Helper function to check if properties match
      const propertiesMatch = (a: any, b: any): boolean => {
        return JSON.stringify(a) === JSON.stringify(b);
      };

      // If we have no current group or properties don't match, create new group
      if (
        !currentGroup ||
        !propertiesMatch(currentGroup.properties, properties)
      ) {
        currentGroup = {
          properties,
          texts: [],
        };
        groups.push(currentGroup);
      }

      // Add text to current group
      currentGroup.texts.push(paragraph.texts.join(''));
    }

    return groups;
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

  static getElementCreationId(
    slideElement: XmlElement,
    stripBrackets?: boolean,
  ): string | undefined {
    const cNvPr = XmlSlideHelper.getNonVisibleProperties(slideElement);
    if (cNvPr) {
      const creationIdElement = cNvPr
        .getElementsByTagName('a16:creationId')
        .item(0);

      if (creationIdElement) {
        const id = creationIdElement.getAttribute('id');
        if (stripBrackets) return id.replace('{', '').replace('}', '');
        return id;
      }
    }
  }

  static getElementNameIdx(slideElement: XmlElement): number {
    const elementName = XmlSlideHelper.getElementName(slideElement);
    if (!elementName) {
      return 0;
    }

    // Find the parent slide element (spTree) to search all elements on the slide
    const currentNode = XmlHelper.getClosestParent('p:spTree', slideElement);

    if (!currentNode) {
      return 0; // Unable to find slide parent
    }

    const spTree = currentNode as XmlElement;

    // Get all named elements from the slide
    const namedElements: XmlElement[] = [];
    const nvPrs = spTree.getElementsByTagNameNS(nsMain, 'cNvPr');

    XmlHelper.modifyCollection(nvPrs, (nvPr: any) => {
      const parentNode = nvPr.parentNode.parentNode;
      const name = nvPr.getAttribute('name');
      if (name === elementName) {
        namedElements.push(parentNode);
      }
    });

    // Find the index of the current element in the array of elements with the same name
    for (let i = 0; i < namedElements.length; i++) {
      if (namedElements[i] === slideElement) {
        return i;
      }
    }

    return 0;
  }

  /**
   * Parses local tag name to specify element type in case it is a 'graphicFrame'.
   * @param slideElementParent
   */
  static getElementType(slideElementParent: XmlElement): ElementType {
    let type = slideElementParent.localName;

    const getUri = () => {
      const graphicData =
        slideElementParent.getElementsByTagName('a:graphicData')[0];
      return graphicData.getAttribute('uri');
    };

    switch (type) {
      case 'graphicFrame':
        type = mapUriType[getUri()] || type;
        break;
      case 'oleObj':
        type = 'OLEObject';
        break;
    }

    // Check for hyperlinks
    const hasHyperlink =
      slideElementParent.getElementsByTagName('a:hlinkClick');
    if (hasHyperlink.length > 0) {
      type = 'Hyperlink';
    }

    return type as ElementType;
  }

  static parseShapeCoordinates(
    slideElementParent: XmlElement,
    returnDefaults?: boolean,
  ): ElementPosition {
    const xFrmsA = slideElementParent.getElementsByTagName('a:xfrm');
    const xFrmsP = slideElementParent.getElementsByTagName('p:xfrm');
    const xFrms = xFrmsP.item(0) ? xFrmsP : xFrmsA;

    const position = {
      x: 0,
      y: 0,
      cx: 0,
      cy: 0,
      rot: 0,
    };

    if (!xFrms.item(0)) {
      if (returnDefaults === false) {
        return null;
      }
      return position;
    }

    const xFrm = xFrms.item(0);

    const Off = xFrm.getElementsByTagName('a:off').item(0);
    const Ext = xFrm.getElementsByTagName('a:ext').item(0);

    position.x = XmlSlideHelper.parseCoordinate(Off, 'x');
    position.y = XmlSlideHelper.parseCoordinate(Off, 'y');
    position.cx = XmlSlideHelper.parseCoordinate(Ext, 'cx');
    position.cy = XmlSlideHelper.parseCoordinate(Ext, 'cy');

    if (xFrm.getAttribute('rot')) {
      position.rot = parseInt(xFrm.getAttribute('rot'));
    }

    return position;
  }

  static parseCoordinate = (
    element: XmlElement,
    attributeName: string,
  ): number => {
    return parseInt(element.getAttribute(attributeName), 10);
  };

  static parseGroupInfo = (element: XmlElement): GroupInfo => {
    // Check if element is a child of a group
    // Look for a parent node that is a group (grpSp)
    const isChild =
      element.parentNode &&
      (element.parentNode.nodeName === 'grpSp' ||
        element.parentNode.nodeName.includes('grpSp'));

    // Check if element is a group parent itself
    const isParent =
      element.localName === 'grpSp' || element.nodeName.includes('grpSp');

    // Function to get the parent group if element is a child
    const getParent = () => {
      if (isChild && element.parentNode) {
        return element.parentNode as XmlElement;
      }
      return null;
    };

    // Function to get children if element is a group parent
    const getChildren = () => {
      if (isParent) {
        // Get all children that are not group properties or group metadata
        const children = Array.from(element.childNodes).filter((node: Node) => {
          if (node.nodeType !== 1) return false; // Skip non-element nodes

          const nodeName =
            (node as Element).localName || (node as Element).nodeName;
          // Skip group property elements
          return (
            !nodeName.includes('nvGrpSpPr') && !nodeName.includes('grpSpPr')
          );
        }) as XmlElement[];

        return children;
      }
      return [];
    };

    return {
      isChild,
      isParent,
      getParent,
      getChildren,
    };
  };

  async getSlideLayoutXml(layoutId: number): Promise<XmlDocument> {
    return XmlSlideHelper.getSlideLayoutXml(this.sourceArchive, layoutId);
  }

  static async getSlideLayoutXml(
    sourceArchive: IArchive,
    layoutId: number,
  ): Promise<XmlDocument> {
    const layoutPath = 'ppt/slideLayouts/slideLayout' + layoutId + '.xml';
    const layoutXml = await XmlHelper.getXmlFromArchive(
      sourceArchive,
      layoutPath,
    );
    if (layoutXml) {
      return layoutXml;
    }
  }

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
      const xml = await XmlHelper.getXmlFromArchive(this.sourceArchive, path);
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

  /**
   * Reconstruct the complete selector for a given xmlElement, including the
   * n-th occurance of the shape name on the current slide
   * @param slideElement
   */
  static getSelector(slideElement: XmlElement): FindElementMultiSelector {
    const creationId = XmlSlideHelper.getElementCreationId(slideElement, true);
    const nameIdx = !creationId
      ? XmlSlideHelper.getElementNameIdx(slideElement)
      : 0;

    return {
      name: XmlSlideHelper.getElementName(slideElement),
      creationId,
      nameIdx,
    };
  }

  /**
   * Determines the type of visual element in PowerPoint
   * @param element The XML element to check
   * @returns A string identifying the element type
   */
  static getElementVisualType(element: XmlElement): ElementVisualType {
    // Check for graphicFrame elements (charts, SmartArt, tables, etc.)
    if (element.tagName === 'p:graphicFrame') {
      const graphicData = XmlHelper.findElement(element, 'a:graphicData');
      if (graphicData) {
        const uri = graphicData.getAttribute('uri').toLowerCase();

        // Check for specific URIs that identify element types
        if (uri && uri.includes('chart')) {
          return 'chart';
        } else if (
          uri &&
          (uri.includes('smartart') || uri.includes('diagram'))
        ) {
          return 'smartArt';
        } else if (uri && uri.includes('table')) {
          return 'table';
        }
      }
      return 'graphicFrame';
    }

    // Check for tables - also check direct table elements
    if (
      XmlHelper.findElement(element, 'a:tbl') ||
      element.getElementsByTagName('a:tbl')[0]
    ) {
      return 'table';
    }

    // Check for 3D models
    if (XmlHelper.findElement(element, 'a:scene3d')) {
      return '3dModel';
    }

    // Check for SVG Images
    if (XmlHelper.findElement(element, 'a:svgBlip')
    || XmlHelper.findElement(element, 'asvg:svgBlip')) {
      return 'svgImage';
    }

    // Check for pictures/photos
    const hasPicPr = !!element.getElementsByTagName('p:nvPicPr')[0];
    if (hasPicPr) {
      return 'picture';
    }

    // Check for image fills
    const hasBlipFill = !!XmlHelper.findElement(element, 'a:blipFill');
    if (hasBlipFill) {
      return 'imageFilledShape';
    }

    const hasGeometry =
      !!XmlHelper.findElement(element, 'a:prstGeom') ||
      !!XmlHelper.findElement(element, 'a:custGeom');
    if (hasGeometry) {
      // Check if it's a line shape specifically
      const prstGeomElement = XmlHelper.findElement(element, 'a:prstGeom');
      if (prstGeomElement) {
        const prst = prstGeomElement.getAttribute('prst');
        // Common line presets in PowerPoint
        const linePresets = [
          'line',
          'lineInv',
          'straightConnector1',
          'bentConnector2',
          'bentConnector3',
          'bentConnector4',
          'bentConnector5',
          'curvedConnector2',
          'curvedConnector3',
          'curvedConnector4',
          'curvedConnector5',
          'callout1',
          'callout2',
          'callout3',
          'accentCallout1',
          'accentCallout2',
          'accentCallout3',
          'borderCallout1',
          'borderCallout2',
          'borderCallout3',
          'accentBorderCallout1',
          'accentBorderCallout2',
          'accentBorderCallout3',
        ];

        if (prst && linePresets.includes(prst)) {
          return 'vectorLine';
        }

        if (prst && prst === 'rect') {
          const cNvSpPr = XmlHelper.findElement(element, 'p:cNvSpPr');
          if(cNvSpPr && cNvSpPr.getAttribute('txBox') === '1' ) {
            return 'textBox'
          }
          return 'rectangle';
        }
      }

      return 'vectorShape';
    }

    // Default case
    return 'unknown';
  }
}
