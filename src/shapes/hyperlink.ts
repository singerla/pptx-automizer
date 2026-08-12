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
import { IShapeAction } from '../interfaces/ishape-action';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import ModifyHyperlinkHelper from '../helper/modify-hyperlink-helper';
import { log } from '../helper/logger';

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const SLIDE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

const basename = (target: string): string => target.split('/').pop();

export class Hyperlink extends Shape implements IShapeAction {
  private hyperlinkType: 'internal' | 'external';
  private hyperlinkTarget: string;
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
    await this.applyCallbacks(this.callbacks, this.targetElement, slideRelXml);

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
    return target.isExternal || target.type === HYPERLINK_REL_TYPE
      ? 'external'
      : 'internal';
  }

  private async editTargetHyperlinkRel(): Promise<void> {
    const isExternalLink = this.hyperlinkType === 'external';
    const rels = await this.getRelationsElement();

    if (this.hyperlinkRelIsUpToDate(rels)) {
      // Rewriting an unchanged relationship would drop the original rId, which
      // can still be in use by other shapes on the same slide.
      return;
    }

    ModifyHyperlinkHelper.setHyperlinkTarget(
      this.hyperlinkTarget,
      isExternalLink,
    )(this.targetElement, rels as any);
  }

  /**
   * Checks whether the relationship referenced by the target element already
   * points to the required hyperlink target.
   */
  private hyperlinkRelIsUpToDate(rels: XmlElement): boolean {
    const hlinkClick = this.targetElement
      ?.getElementsByTagName('a:hlinkClick')
      .item(0);
    const rId = hlinkClick?.getAttribute('r:id');
    if (!rId || !this.hyperlinkTarget) {
      return false;
    }

    const existingRel = XmlHelper.findByAttributeValue(
      rels.getElementsByTagName('Relationship'),
      'Id',
      rId,
    )[0];
    if (!existingRel) {
      return false;
    }

    const isExternalLink = this.hyperlinkType === 'external';
    if (existingRel.getAttribute('Type') !== this.relationTypeUrl()) {
      return false;
    }

    const existingTarget = existingRel.getAttribute('Target') || '';
    return isExternalLink
      ? existingTarget === this.hyperlinkTarget
      : // Internal targets are stored either as `slide2.xml` or as
        // `../slides/slide2.xml`, both resolving to the same slide.
        basename(existingTarget) === basename(this.hyperlinkTarget);
  }

  private relationTypeUrl(): string {
    return this.hyperlinkType === 'external'
      ? HYPERLINK_REL_TYPE
      : SLIDE_REL_TYPE;
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    return XmlHelper.getRelationshipItems(
      archive,
      relsPath,
      (element: XmlElement, rels: Target[]) => {
        const type = element.getAttribute('Type');
        if (type === HYPERLINK_REL_TYPE || type === SLIDE_REL_TYPE) {
          rels.push({
            rId: element.getAttribute('Id'),
            type: element.getAttribute('Type'),
            file: element.getAttribute('Target'),
            filename: element.getAttribute('Target'),
            element: element,
            isExternal:
              element.getAttribute('TargetMode') === 'External' ||
              type === HYPERLINK_REL_TYPE,
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
      log.debug(
        'modifyOnAddedSlide called on Hyperlink without a valid source target/rId.',
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
