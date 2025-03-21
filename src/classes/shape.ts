import { XmlHelper } from '../helper/xml-helper';
import { GeneralHelper, vd } from '../helper/general-helper';
import {
  ChartModificationCallback,
  ImportedElement,
  ShapeModificationCallback,
  ShapeTargetType,
  Target,
  Workbook,
} from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlDocument, XmlElement } from '../types/xml-types';
import {
  ContentTypeExtension,
  ContentTypeMap,
} from '../enums/content-type-map';
import { ElementSubtype } from '../enums/element-type';
import IArchive from '../interfaces/iarchive';

export class Shape {
  mode: string;
  name: string;

  sourceArchive: IArchive;
  sourceSlideNumber: number;
  sourceSlideFile: string;
  sourceNumber: number;
  sourceFile: string;
  sourceRid: string;
  sourceElement: XmlElement;

  targetFile: string;
  targetArchive: IArchive;
  targetTemplate: RootPresTemplate;
  targetSlideNumber: number;
  targetNumber: number;
  targetSlideFile: string;
  targetSlideRelFile: string;

  createdRid: string;

  relRootTag: string;
  relAttribute: string;
  relType: string;
  relParent: (element: XmlElement) => XmlElement;

  targetElement: XmlElement;
  targetType: ShapeTargetType;
  target: Target;

  callbacks: (ShapeModificationCallback | ChartModificationCallback)[];
  hasCreationId: boolean;
  contentTypeMap: typeof ContentTypeMap;
  subtype: ElementSubtype;
  shape: ImportedElement;

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    this.shape = shape;

    this.mode = shape.mode;
    this.name = shape.name;
    this.targetType = targetType;

    this.sourceArchive = shape.sourceArchive;
    this.sourceSlideNumber = shape.sourceSlideNumber;
    this.sourceSlideFile = `ppt/slides/slide${this.sourceSlideNumber}.xml`;
    this.sourceElement = shape.sourceElement;
    this.hasCreationId = shape.hasCreationId;

    this.callbacks = GeneralHelper.arrayify(shape.callback);
    this.contentTypeMap = ContentTypeMap;

    if (shape.target) {
      this.sourceNumber = shape.target.number;
      this.sourceRid = shape.target.rId;
      this.subtype = shape.target.subtype;
      this.target = shape.target;
    }
  }

  async setTarget(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    const targetType = this.targetType;

    this.targetTemplate = targetTemplate;
    this.targetArchive = await this.targetTemplate.archive;
    this.targetSlideNumber = targetSlideNumber;
    this.targetSlideFile = `ppt/${targetType}s/${targetType}${this.targetSlideNumber}.xml`;
    this.targetSlideRelFile = `ppt/${targetType}s/_rels/${targetType}${this.targetSlideNumber}.xml.rels`;
  }

  async setTargetElement(): Promise<void> {
    if (!this.sourceElement) {
      // If we don't have a source element, we might be trying to remove a hyperlink
      console.log(
        `Warning: No source element for shape ${this.name}. Creating empty element for operations.`,
      );
      if (this.shape && this.shape.mode === 'remove' && this.targetArchive) {
        // For remove operations, we don't need a source element
        // Just continue without setting targetElement
        return;
      }

      // For non-remove operations or if other conditions aren't met, throw the error
      console.log(this.shape);
      throw `No source element for shape ${this.name}`;
    }
    this.targetElement = this.sourceElement.cloneNode(true) as XmlElement;
  }

  async appendToSlideTree(): Promise<void> {
    const targetSlideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideFile,
    );

    targetSlideXml
      .getElementsByTagName('p:spTree')[0]
      .appendChild(this.targetElement);

    // Process hyperlinks in the element if this is a hyperlink element
    if (this.relRootTag === 'a:hlinkClick') {
      await this.processHyperlinks(targetSlideXml);
    }

    XmlHelper.writeXmlToArchive(
      this.targetArchive,
      this.targetSlideFile,
      targetSlideXml,
    );
  }

  // Process hyperlinks in the element
  async processHyperlinks(targetSlideXml: XmlDocument): Promise<void> {
    // Find all text runs in the element
    const runs = this.targetElement.getElementsByTagName('a:r');

    for (let i = 0; i < runs.length; i++) {
      const run = runs[i];
      const rPr = run.getElementsByTagName('a:rPr')[0];

      if (rPr) {
        // Find hyperlink elements
        const hlinkClicks = rPr.getElementsByTagName('a:hlinkClick');

        for (let j = 0; j < hlinkClicks.length; j++) {
          const hlinkClick = hlinkClicks[j];
          const sourceRid = hlinkClick.getAttribute('r:id');

          if (sourceRid) {
            // Update the r:id attribute to use the created relationship ID
            hlinkClick.setAttribute('r:id', this.createdRid);

            // Ensure the xmlns:r attribute is set
            hlinkClick.setAttribute(
              'xmlns:r',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            );
          }
        }
      }
    }
  }

  async replaceIntoSlideTree(): Promise<void> {
    await this.modifySlideTree(true);
  }

  async removeFromSlideTree(): Promise<void> {
    await this.modifySlideTree(false);
  }

  async modifySlideTree(insertBefore?: boolean): Promise<void> {
    const archive = this.targetArchive;
    const slideFile = this.targetSlideFile;

    const targetSlideXml = await XmlHelper.getXmlFromArchive(
      archive,
      slideFile,
    );

    const findMethod = this.hasCreationId ? 'findByCreationId' : 'findByName';

    const sourceElementOnTargetSlide = await XmlHelper[findMethod](
      targetSlideXml,
      this.name,
    );

    if (!sourceElementOnTargetSlide?.parentNode) {
      console.error(`Can't modify slide tree for ${this.name}`);
      return;
    }

    if (insertBefore === true) {
      sourceElementOnTargetSlide.parentNode.insertBefore(
        this.targetElement,
        sourceElementOnTargetSlide,
      );
    }

    sourceElementOnTargetSlide.parentNode.removeChild(
      sourceElementOnTargetSlide,
    );

    // Process hyperlinks in the element if this is a hyperlink element
    if (this.relRootTag === 'a:hlinkClick') {
      await this.processHyperlinks(targetSlideXml);
    }

    XmlHelper.writeXmlToArchive(archive, slideFile, targetSlideXml);
  }

  async updateElementsRelId(): Promise<void> {
    const targetSlideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideFile,
    );
    const targetElements = await this.getElementsByRid(
      targetSlideXml,
      this.sourceRid,
    );

    targetElements.forEach((targetElement: XmlElement) => {
      this.relParent(targetElement)
        .getElementsByTagName(this.relRootTag)[0]
        .setAttribute(this.relAttribute, this.createdRid);
    });

    XmlHelper.writeXmlToArchive(
      this.targetArchive,
      this.targetSlideFile,
      targetSlideXml,
    );
  }

  /*
   * This will find all elements with a matching rId on a
   * <p:cSld>, including related images at <p:bg> and <p:spTree>.
   */
  async getElementsByRid(
    slideXml: XmlDocument,
    rId: string,
  ): Promise<XmlElement[]> {
    const sourceList = slideXml
      .getElementsByTagName('p:cSld')[0]
      .getElementsByTagName(this.relRootTag);

    const sourceElements = XmlHelper.findByAttributeValue(
      sourceList,
      this.relAttribute,
      rId,
    );

    return sourceElements;
  }

  async updateTargetElementRelId(): Promise<void> {
    this.targetElement
      .getElementsByTagName(this.relRootTag)
      .item(0)
      .setAttribute(this.relAttribute, this.createdRid);
  }

  applyCallbacks(
    callbacks: ShapeModificationCallback[],
    element: XmlElement,
    relation?: XmlElement,
  ): void {
    callbacks.forEach((callback) => {
      if (typeof callback === 'function') {
        try {
          callback(element, relation);
        } catch (e) {
          console.warn(e);
        }
      }
    });
  }

  applyChartCallbacks(
    callbacks: ChartModificationCallback[],
    element: XmlElement,
    chart: XmlDocument,
    workbook: Workbook,
  ): void {
    callbacks.forEach((callback) => {
      if (typeof callback === 'function') {
        try {
          callback(element, chart, workbook);
        } catch (e) {
          console.warn(e);
        }
      }
    });
  }

  appendImageExtensionToContentType(
    extension: ContentTypeExtension,
  ): Promise<XmlElement | boolean> {
    return XmlHelper.appendImageExtensionToContentType(
      this.targetArchive,
      extension,
    );
  }
}
