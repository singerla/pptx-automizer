import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';
import { RelationshipAttribute, XmlElement } from '../types/xml-types';
import {
  ImportedElement,
  ShapeModificationCallback,
  ShapeTargetType,
  Target,
} from '../types/types';
import { IImage } from '../interfaces/iimage';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ElementType } from '../enums/element-type';
import IArchive from '../interfaces/iarchive';
import { ContentTypeExtension } from '../enums/content-type-map';
import { TargetByRelIdMap } from '../constants/constants';

export class Image extends Shape implements IImage {
  extension: ContentTypeExtension;
  createdRelation: XmlElement;
  callbacks: ShapeModificationCallback[];

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    super(shape, targetType);

    this.sourceFile = shape.target.file.replace('../media/', '');
    this.extension = FileHelper.getFileExtension(this.sourceFile);
    this.relAttribute = 'r:embed';
    this.relType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

    // A shape retrieved by Image.getAllOnSlide() can also be (nested) SVG.
    if (!shape.sourceMode && this.extension === 'svg') {
      shape.sourceMode = 'image:svg';
    }

    switch (shape.sourceMode) {
      case 'image:svg':
        this.relRootTag = TargetByRelIdMap['image:svg'].relRootTag;
        this.relParent = (element: XmlElement) =>
          element.parentNode as XmlElement;
        break;
      case 'image:media':
      case 'image:audioFile':
      case 'image:videoFile':
        this.relRootTag = TargetByRelIdMap[shape.sourceMode].relRootTag;
        this.relAttribute = TargetByRelIdMap[shape.sourceMode].relAttribute;
        this.relType = TargetByRelIdMap[shape.sourceMode].relType;
        this.relParent = (element: XmlElement) =>
          element.parentNode as XmlElement;
        break;
      default:
        this.relRootTag = 'a:blip';
        this.relParent = (element: XmlElement) =>
          element.parentNode.parentNode as XmlElement;
        break;
    }
  }

  /*
   * It is necessary to update existing rIds for all
   * unmodified images on an added slide at first.
   */
  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.updateElementsRelId();

    return this;
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);

    await this.setTargetElement();
    await this.updateTargetElementRelId();

    this.applyImageCallbacks();

    await this.replaceIntoSlideTree();

    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.setTargetElement();

    await this.updateTargetElementRelId();
    await this.appendToSlideTree();

    this.applyImageCallbacks();

    await this.processImageRelations(targetTemplate, targetSlideNumber);

    return this;
  }

  /**
   * For audio/video and svg, some more relations need to be handled.
   */
  async processImageRelations(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ) {
    /*
     * SVG images require a corresponding PNG image.
     */
    if (this.hasSvgBlipRelation()) {
      await this.processRelatedContent(
        targetTemplate,
        targetSlideNumber,
        'image:svg',
      );
    }

    /**
     * Media files are children of images with additional relations
     */
    if (this.hasAudioRelation()) {
      await this.processRelatedMediaContent(
        targetTemplate,
        targetSlideNumber,
        'image:audioFile',
      );
    }
    if (this.hasVideoRelation()) {
      await this.processRelatedMediaContent(
        targetTemplate,
        targetSlideNumber,
        'image:videoFile',
      );
    }
  }

  async processRelatedMediaContent(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    sourceMode: ImportedElement['sourceMode'],
  ) {
    await this.processRelatedContent(
      targetTemplate,
      targetSlideNumber,
      'image:media',
    );
    await this.processRelatedContent(
      targetTemplate,
      targetSlideNumber,
      sourceMode,
    );
  }

  async processRelatedContent(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    sourceMode: ImportedElement['sourceMode'],
  ) {
    const relsPath = `ppt/slides/_rels/slide${this.sourceSlideNumber}.xml.rels`;

    const target = await XmlHelper.getTargetByRelId(
      this.sourceArchive,
      relsPath,
      this.targetElement,
      sourceMode,
    );
    await new Image(
      {
        mode: 'append',
        target,
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceSlideNumber,
        type: ElementType.Image,
        sourceMode,
      },
      this.targetType,
    ).modifyMediaRelation(
      targetTemplate,
      targetSlideNumber,
      this.targetElement,
    );
  }

  async modifyMediaRelation(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    targetElement: XmlElement,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);

    this.targetElement = targetElement;
    await this.updateTargetElementRelId();

    return this;
  }

  /*
   * Apply all ShapeModificationCallbacks to target element.
   * Third argument this.createdRelation is necessery to directly
   * manipulate relation Target and change the image.
   */
  applyImageCallbacks() {
    this.applyCallbacks(
      this.callbacks,
      this.targetElement,
      this.createdRelation,
    );
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    this.targetNumber = this.targetTemplate.incrementCounter('images');
    this.targetFile = this.getTargetFileName();

    await this.copyFiles();
    await this.appendTypes();
    await this.appendToSlideRels();
  }

  async copyFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/media/${this.sourceFile}`,
      this.targetArchive,
      `ppt/media/${this.targetFile}`,
    );
  }

  getTargetFileName(): string {
    const targetFileType = this.target.file.includes('media')
      ? 'media'
      : 'image';
    return targetFileType + this.targetNumber + '.' + this.extension;
  }

  async appendTypes(): Promise<void> {
    await this.appendImageExtensionToContentType(this.extension);
  }

  /**
   * ToDo: This will always append a new relation, and never replace an
   * existing relation. At the end of creation process, unused relations will
   * remain existing in the .xml.rels file. PowerPoint will not complain, but
   * integrity checks will not be valid by this.
   */
  async appendToSlideRels(): Promise<void> {
    const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
    this.createdRid = await XmlHelper.getNextRelId(
      this.targetArchive,
      targetRelFile,
    );

    const targetFileName = this.getTargetFileName();

    const attributes = {
      Id: this.createdRid,
      Type: this.relType,
      Target: `../media/${targetFileName}`,
    } as RelationshipAttribute;

    this.createdRelation = await XmlHelper.append(
      XmlHelper.createRelationshipChild(
        this.targetArchive,
        targetRelFile,
        attributes,
      ),
    );
  }

  hasSvgBlipRelation(): boolean {
    return this.targetElement.getElementsByTagName('asvg:svgBlip').length > 0;
  }

  hasAudioRelation(): boolean {
    return this.targetElement.getElementsByTagName('a:audioFile').length > 0;
  }

  hasVideoRelation(): boolean {
    return this.targetElement.getElementsByTagName('a:videoFile').length > 0;
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    return await XmlHelper.getTargetsByRelationshipType(
      archive,
      relsPath,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    );
  }
}
