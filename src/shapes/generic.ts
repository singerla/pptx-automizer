import {
  ImportedElement,
  ShapeModificationCallback,
  ShapeTargetType,
} from '../types/types';
import { IShapeAction } from '../interfaces/ishape-action';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { Shape } from '../classes/shape';
import { XmlElement } from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { HyperlinkProcessor } from '../helper/hyperlink-processor';
import { Image } from './image';
import { ElementType } from '../enums/element-type';
import { PptPaths } from '../helper/ppt-paths';
import { log } from '../helper/logger';

export class GenericShape extends Shape implements IShapeAction {
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

    // If this element contains images, copy the media files and relationships
    await this.copyImageRelationships(targetTemplate, targetSlideNumber);

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
    await this.applyCallbacks(this.callbacks, this.targetElement, slideRelXml.documentElement as XmlElement);
  }

  /**
   * Copy hyperlink relationships from source slide to target slide
   */
  async copyHyperlinkRelationships(_targetSlideNumber: number): Promise<void> {
    if (!this.targetElement) return;

    await HyperlinkProcessor.copyMultipleHyperlinks(
      this.targetElement,
      this.sourceArchive,
      this.sourceSlideNumber,
      this.targetArchive,
      this.targetSlideRelFile,
      // The unmutated source element separates imported hyperlink ids (to
      // resolve against the source rels) from ids added by modification
      // callbacks during prepare(), whose rels already live on the target.
      this.sourceElement,
    );
  }

  /**
   * Copy image relations from source slide to target slide.
   *
   * A generic shape can be filled with an image (`a:blipFill` inside
   * `p:spPr`), and a group shape can contain any number of pictures. None of
   * these are `p:pic` elements, so the Image shape class never sees them and
   * their media files and relations would be missing on the target slide.
   */
  async copyImageRelationships(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    if (!this.targetElement) return;

    const blips = [
      ...Array.from(this.targetElement.getElementsByTagName('a:blip')),
      // An svg blip is nested inside the a:blip holding its png fallback and
      // requires a relation of its own.
      ...Array.from(this.targetElement.getElementsByTagName('asvg:svgBlip')),
    ].filter((blip) => blip.getAttribute('r:embed'));

    if (!blips.length) return;

    const sourceImages = await Image.getAllOnSlide(
      this.sourceArchive,
      PptPaths.slideRels(this.sourceSlideNumber),
    );

    for (const blip of blips) {
      const sourceRid = blip.getAttribute('r:embed');
      const target = sourceImages.find((image) => image.rId === sourceRid);

      if (!target) {
        log.warn(
          `Can't find image relation ${sourceRid} of ${this.name} on slide ${this.sourceSlideNumber}`,
        );
        continue;
      }

      const image = new Image(
        {
          mode: 'append',
          target,
          sourceArchive: this.sourceArchive,
          sourceSlideNumber: this.sourceSlideNumber,
          type: ElementType.Image,
        },
        this.targetType,
      );

      // Copies the media file, registers its content type and appends a
      // relation to the target slide rels.
      await image.prepare(targetTemplate, targetSlideNumber);

      blip.setAttribute('r:embed', image.createdRid);
    }
  }
}
