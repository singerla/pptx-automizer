import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';
import {
  HelperElement,
  RelationshipAttribute,
  XmlElement,
} from '../types/xml-types';
import { ImportedElement, ShapeTargetType, Target } from '../types/types';
import { IImage } from '../interfaces/iimage';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ElementType } from '../enums/element-type';
import IArchive from '../interfaces/iarchive';

export class Image extends Shape implements IImage {
  extension: string;

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

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.updateElementsRelId();

    return this;
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.updateElementsRelId();

    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.setTargetElement();

    this.applyCallbacks(this.callbacks, this.targetElement);

    await this.updateTargetElementRelId();
    await this.appendToSlideTree();

    if (this.hasSvgRelation()) {
      const target = await XmlHelper.getTargetByRelId(
        this.sourceArchive,
        this.sourceSlideNumber,
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
      ).modify(targetTemplate, targetSlideNumber);
    }

    return this;
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
  async appendToSlideRels(): Promise<HelperElement> {
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

    return XmlHelper.append(
      XmlHelper.createRelationshipChild(
        this.targetArchive,
        targetRelFile,
        attributes,
      ),
    );
  }

  hasSvgRelation(): boolean {
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
