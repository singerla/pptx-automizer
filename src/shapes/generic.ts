import {
  ImportedElement,
  ShapeModificationCallback,
  ShapeTargetType,
} from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { Shape } from '../classes/shape';
import { XmlElement } from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { HyperlinkProcessor } from '../helper/hyperlink-processor';
import { vd } from '../helper/general-helper';

export class GenericShape extends Shape {
  sourceElement: XmlElement;
  callbacks: ShapeModificationCallback[];

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    super(shape, targetType);
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.replaceIntoSlideTree();
    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.appendToSlideTree();

    // If this element contains hyperlinks, copy the hyperlink relationships
    await this.copyHyperlinkRelationships(targetSlideNumber);

    return this;
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<GenericShape> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);
    await this.setTargetElement();

    // Get the slide relations XML to pass to callbacks
    const slideRelXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideRelFile
    );

    // Pass both the element and the relation to applyCallbacks
    // Use the documentElement property to get the root element of the XML document
    this.applyCallbacks(this.callbacks, this.targetElement, slideRelXml.documentElement as XmlElement);
  }

  /**
   * Copy hyperlink relationships from source slide to target slide
   */
  async copyHyperlinkRelationships(targetSlideNumber: number): Promise<void> {
    if (!this.targetElement) return;

    await HyperlinkProcessor.copyMultipleHyperlinks(
      this.targetElement,
      this.sourceArchive,
      this.sourceSlideNumber,
      this.targetArchive,
      this.targetSlideRelFile
    );
  }
}
