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

export class Image extends Shape implements IImage {
  extension: ContentTypeExtension;
  createdRelation: XmlElement;
  callbacks: ShapeModificationCallback[];

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    super(shape, targetType);

    this.sourceFile = shape.target.file.replace('../media/', '');
    this.extension = FileHelper.getFileExtension(this.sourceFile);
    this.relAttribute = 'r:embed';

    switch (this.extension) {
      case 'svg':
        this.relRootTag = 'asvg:svgBlip';
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

  async modifySvgRelation(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    targetElement: XmlElement,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);

    this.targetElement = targetElement;
    await this.updateTargetElementRelId();

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

    /*
     * SVG images require a corresponding PNG image.
     */
    if (this.hasSvgBlipRelation()) {
      const relsPath = `ppt/slides/_rels/slide${this.sourceSlideNumber}.xml.rels`;
      const target = await XmlHelper.getTargetByRelId(
        this.sourceArchive,
        relsPath,
        this.targetElement,
        'image:svg',
      );
      await new Image(
        {
          mode: 'append',
          target,
          sourceArchive: this.sourceArchive,
          sourceSlideNumber: this.sourceSlideNumber,
          type: ElementType.Image,
        },
        this.targetType,
      ).modifySvgRelation(
        targetTemplate,
        targetSlideNumber,
        this.targetElement,
      );
    }

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
    this.targetFile = 'image' + this.targetNumber + '.' + this.extension;

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

    const attributes = {
      Id: this.createdRid,
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      Target: `../media/image${this.targetNumber}.${this.extension}`,
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
