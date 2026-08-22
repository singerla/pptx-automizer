import { XmlHelper } from '../helper/xml-helper';
import { GeneralHelper } from '../helper/general-helper';
import { PptPaths } from '../helper/ppt-paths';
import { log } from '../helper/logger';
import { CallbackError } from '../errors';
import { HyperlinkProcessor } from '../helper/hyperlink-processor';
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
    this.sourceSlideFile = PptPaths.slide(this.sourceSlideNumber);
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
    this.targetArchive = this.targetTemplate.archive;
    this.targetSlideNumber = targetSlideNumber;
    this.targetSlideFile = PptPaths.part(targetType, this.targetSlideNumber);
    this.targetSlideRelFile = PptPaths.partRels(
      targetType,
      this.targetSlideNumber,
    );
  }

  async setTargetElement(): Promise<void> {
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

    // A hyperlink shape cloned from a source slide still carries its source
    // r:id, which may collide with an unrelated relationship already declared
    // on the target slide. Rewrite it to the reserved createdRid — guaranteed
    // unused — before Hyperlink.append() creates the matching relationship.
    if (this.relRootTag === 'a:hlinkClick') {
      await this.processHyperlinks();
    }

    XmlHelper.writeXmlToArchive(
      this.targetArchive,
      this.targetSlideFile,
      targetSlideXml,
    );
  }

  /**
   * Rewrites the single a:hlinkClick r:id of this element to createdRid.
   * Only used on append: on modify, Hyperlink.modify() manages the existing
   * relationship in place via editTargetHyperlinkRel(), and overwriting its
   * r:id here would leave it without a matching relationship entry.
   */
  async processHyperlinks(): Promise<void> {
    if (!this.targetElement || !this.createdRid) return;

    await HyperlinkProcessor.processSingleHyperlink(
      this.targetElement,
      this.createdRid,
    );
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
    const selector = this.hasCreationId
      ? this.name
      : this.shape.selector.name;

    const sourceElementOnTargetSlide = await XmlHelper[findMethod](
      targetSlideXml,
      selector,
      this.shape.selector.nameIdx,
    );

    if (!sourceElementOnTargetSlide?.parentNode) {
      log.error(`Can't modify slide tree for ${this.name}`);
      return;
    }

    if (insertBefore === true && this.targetElement) {
      sourceElementOnTargetSlide.parentNode.insertBefore(
        this.targetElement,
        sourceElementOnTargetSlide,
      );
    }

    sourceElementOnTargetSlide.parentNode.removeChild(
      sourceElementOnTargetSlide,
    );

    XmlHelper.writeXmlToArchive(archive, slideFile, targetSlideXml);
  }

  async updateElementsRelId(
    cb?: (targetElement: XmlElement) => void,
  ): Promise<void> {
    const targetSlideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideFile,
    );

    const targetElements = await this.getElementsByRid(
      targetSlideXml,
      this.sourceRid,
    );

    targetElements.forEach((targetElement: XmlElement) => {
      if (cb && typeof cb === 'function') {
        cb(targetElement);
      } else {
        this.relParent(targetElement)
          .getElementsByTagName(this.relRootTag)[0]
          .setAttribute(this.relAttribute, this.createdRid);
      }
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

    return XmlHelper.findByAttributeValue(sourceList, this.relAttribute, rId);
  }

  async updateTargetElementRelId(): Promise<void> {
    this.targetElement
      .getElementsByTagName(this.relRootTag)
      .item(0)
      .setAttribute(this.relAttribute, this.createdRid);
  }

  async applyCallbacks(
    callbacks: ShapeModificationCallback[],
    element: XmlElement,
    relation?: XmlElement,
  ): Promise<void> {
    for (const callback of callbacks) {
      if (typeof callback === 'function') {
        try {
          await callback(element, relation);
        } catch (e) {
          this.handleCallbackError(e);
        }
      }
    }
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
          this.handleCallbackError(e);
        }
      }
    });
  }

  /**
   * A throwing modification callback fails the run by default.
   * With `AutomizerParams.continueOnError`, it is logged and skipped.
   */
  private handleCallbackError(e: unknown): void {
    if (this.shape.continueOnError === true) {
      log.warn(
        `Modification callback failed on element "${this.name}" (${this.targetSlideFile}):`,
        e,
      );
      return;
    }
    throw new CallbackError(this.name, this.targetSlideFile, e);
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
