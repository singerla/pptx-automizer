import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';
import {
  ImportedElement,
  ShapeModificationCallback,
  ShapeTargetType,
  Target,
} from '../types/types';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import ModifyHyperlinkHelper from '../helper/modify-hyperlink-helper';
import { Logger } from '../helper/general-helper';
import { HyperlinkInfo } from '../types/modify-types';

export class Hyperlink extends Shape {
  private hyperlinkType: HyperlinkInfo['type'];
  private hyperlinkTarget: HyperlinkInfo['target'];
  callbacks: ShapeModificationCallback[];

  constructor(
    shape: ImportedElement,
    targetType: ShapeTargetType,
    sourceArchive: IArchive,
    hyperlinkType: 'internal' | 'external' = 'external',
    hyperlinkTarget: string,
  ) {
    super(shape, targetType);
    this.sourceArchive = sourceArchive;
    this.hyperlinkType = hyperlinkType;
    this.hyperlinkTarget = hyperlinkTarget || '';
    this.relRootTag = 'a:hlinkClick';
    this.relAttribute = 'r:id';
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Hyperlink> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.setTargetElement();
    await this.editTargetHyperlinkRel();
    await this.replaceIntoSlideTree();

    // Get the slide relations XML to pass to callbacks
    const slideRelXml = await this.getRelationsElement();

    // Pass both the element and the relation to applyCallbacks
    // Use the documentElement property to get the root element of the XML document
    this.applyCallbacks(this.callbacks, this.targetElement, slideRelXml);

    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Hyperlink> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.setTargetElement();
    await this.appendToSlideTree();

    const slideRelXml = await this.getRelationsElement();
    ModifyHyperlinkHelper.addHyperlink(
      this.hyperlinkTarget,
      this.hyperlinkType === 'internal',
    )(this.targetElement, slideRelXml);

    return this;
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Hyperlink> {
    await this.prepare(targetTemplate, targetSlideNumber);

    if (this.target && this.target.rId) {
      this.sourceRid = this.target.rId;
    }
    const slideRelXml = await this.getRelationsElement();
    ModifyHyperlinkHelper.removeHyperlink()(this.targetElement, slideRelXml);
    await this.removeFromSlideTree();

    return this;
  }

  private async getRelationsElement(): Promise<XmlElement> {
    const slideRelXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideRelFile,
    );
    return slideRelXml.documentElement;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    if (!this.createdRid) {
      const baseId = await XmlHelper.getNextRelId(
        this.targetArchive,
        this.targetSlideRelFile,
      );
      this.createdRid = baseId.endsWith('-created')
        ? baseId.slice(0, -8)
        : baseId;
    }
    if (this.shape && this.shape.target && this.shape.target.rId) {
      this.sourceRid = this.shape.target.rId;
    }
    if (
      !this.hyperlinkTarget &&
      this.shape &&
      this.shape.target &&
      this.shape.target.file
    ) {
      this.hyperlinkTarget = this.shape.target.file;
      this.hyperlinkType = this.determineHyperlinkType(this.shape.target);
    }
  }

  private determineHyperlinkType(target: Target): 'internal' | 'external' {
    return target.isExternal ||
      target.type ===
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
      ? 'external'
      : 'internal';
  }

  private async editTargetHyperlinkRel(): Promise<void> {
    const isExternalLink = this.hyperlinkType === 'external';
    const rels = await this.getRelationsElement();

    ModifyHyperlinkHelper.setHyperlinkTarget(
      this.hyperlinkTarget,
      isExternalLink,
    )(this.targetElement, rels as any);
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    const hyperlinkRelType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    const slideRelType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
    return XmlHelper.getRelationshipItems(
      archive,
      relsPath,
      (element: XmlElement, rels: Target[]) => {
        const type = element.getAttribute('Type');
        if (type === hyperlinkRelType || type === slideRelType) {
          rels.push({
            rId: element.getAttribute('Id'),
            type: element.getAttribute('Type'),
            file: element.getAttribute('Target'),
            filename: element.getAttribute('Target'),
            element: element,
            isExternal:
              element.getAttribute('TargetMode') === 'External' ||
              type === hyperlinkRelType,
          } as Target);
        }
      },
    );
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    if (!this.target || !this.target.rId) {
      Logger.log(
        'modifyOnAddedSlide called on Hyperlink without a valid source target/rId.',
        2,
      );
      return;
    }

    this.sourceRid = this.target.rId;
    this.hyperlinkTarget = this.target.file;
    this.hyperlinkType = this.determineHyperlinkType(this.target);

    await this.prepare(targetTemplate, targetSlideNumber);
    await this.editTargetHyperlinkRel();
  }

  static async addHyperlinkToShape(
    archive: IArchive,
    slidePath: string,
    slideRelsPath: string,
    shapeId: string,
    hyperlinkTarget: string | number,
  ): Promise<string> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, slidePath);
    const shape = XmlHelper.isElementCreationId(shapeId)
      ? XmlHelper.findByCreationId(slideXml, shapeId)
      : XmlHelper.findByName(slideXml, shapeId);

    if (!shape) {
      throw new Error(`Shape with ID/name "${shapeId}" not found`);
    }

    const relXml = await XmlHelper.getXmlFromArchive(archive, slideRelsPath);

    ModifyHyperlinkHelper.addHyperlink(
      hyperlinkTarget,
      typeof hyperlinkTarget === 'number',
    )(shape, relXml.firstChild as XmlElement);

    XmlHelper.writeXmlToArchive(archive, slideRelsPath, relXml);
    XmlHelper.writeXmlToArchive(archive, slidePath, slideXml);

    return await XmlHelper.getNextRelId(archive, slideRelsPath);
  }
}
