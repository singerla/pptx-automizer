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
import { RootPresTemplate } from '../interfaces/root-pres-template';
import IArchive from '../interfaces/iarchive';
import { TargetByRelIdMap } from '../constants/constants';

export class Diagram extends Shape {
  sourceElement: XmlElement;
  relTypeColors: string;
  relTypeData: string;
  relTypeLayout: string;
  relTypeQuickStyle: string;
  relTypeDrawing: string;
  callbacks: ShapeModificationCallback[];
  createdRids: Record<string, string> = {};

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    super(shape, targetType);

    this.relRootTag = TargetByRelIdMap.diagram.relRootTag;
    this.relAttribute = TargetByRelIdMap.diagram.relAttribute;
    this.relParent = (element: XmlElement) =>
      element.parentNode.parentNode.parentNode as XmlElement;

    this.relTypeData =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData';
    this.relTypeColors =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors';
    this.relTypeLayout =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout';
    this.relTypeQuickStyle =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle';
    this.relTypeDrawing =
      'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Diagram> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.clone();

    await this.replaceIntoSlideTree();
    await this.updateRelIds();

    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Diagram> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.clone();

    await this.appendToSlideTree();
    await this.updateRelIds();

    return this;
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Diagram> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();

    return this;
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Diagram> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.updateRelIds();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);
    this.targetNumber = this.targetTemplate.incrementCounter('diagrams');

    await this.copyFiles();
    await this.appendTypes();

    await this.appendToSlideRels(
      this.relTypeData,
      `../diagrams/data${this.targetNumber}.xml`,
    );
    await this.appendToSlideRels(
      this.relTypeColors,
      `../diagrams/colors${this.targetNumber}.xml`,
    );
    await this.appendToSlideRels(
      this.relTypeLayout,
      `../diagrams/layout${this.targetNumber}.xml`,
    );
    await this.appendToSlideRels(
      this.relTypeQuickStyle,
      `../diagrams/quickStyle${this.targetNumber}.xml`,
    );

    // drawing xml will be copied, but it has no explicit relation attribute in slide.xml
    await this.appendToSlideRels(
      this.relTypeDrawing,
      `../diagrams/drawing${this.targetNumber}.xml`,
    );
  }

  async clone(): Promise<void> {
    await this.setTargetElement();
    this.applyCallbacks(this.callbacks, this.targetElement);
  }

  async updateRelIds() {
    await this.updateElementsRelId((targetElement: XmlElement) => {
      targetElement.setAttribute('r:dm', this.createdRids[this.relTypeData]);
      targetElement.setAttribute('r:lo', this.createdRids[this.relTypeLayout]);
      targetElement.setAttribute(
        'r:qs',
        this.createdRids[this.relTypeQuickStyle],
      );
      targetElement.setAttribute('r:cs', this.createdRids[this.relTypeColors]);
    });
  }

  async copyFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/diagrams/data${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/diagrams/data${this.targetNumber}.xml`,
    );
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/diagrams/colors${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/diagrams/colors${this.targetNumber}.xml`,
    );
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/diagrams/drawing${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/diagrams/drawing${this.targetNumber}.xml`,
    );
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/diagrams/layout${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/diagrams/layout${this.targetNumber}.xml`,
    );
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/diagrams/quickStyle${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/diagrams/quickStyle${this.targetNumber}.xml`,
    );
  }

  async appendTypes(): Promise<void> {
    await this.appendDataContentType();
    await this.appendColorsToContentType();
    await this.appendLyoutToContentType();
    await this.appendQuickStyleToContentType();
    await this.appendDrawingToContentType();
  }

  async appendToSlideRels(type: string, target: string): Promise<XmlElement> {
    this.createdRid = await XmlHelper.getNextRelId(
      this.targetArchive,
      this.targetSlideRelFile,
    );

    this.createdRids[type] = this.createdRid;

    const attributes = {
      Id: this.createdRid,
      Type: type,
      Target: target,
    } as RelationshipAttribute;

    return XmlHelper.append(
      XmlHelper.createRelationshipChild(
        this.targetArchive,
        this.targetSlideRelFile,
        attributes,
      ),
    );
  }

  appendDataContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/diagrams/data${this.targetNumber}.xml`,
        ContentType:
          'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
      }),
    );
  }

  appendColorsToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/diagrams/colors${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml`,
      }),
    );
  }

  appendLyoutToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/diagrams/layout${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml`,
      }),
    );
  }

  appendQuickStyleToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/diagrams/quickStyle${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.diagramQuickStyle+xml`,
      }),
    );
  }

  appendDrawingToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/diagrams/drawing${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.diagramDrawing+xml`,
      }),
    );
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    return await XmlHelper.getRelationshipTargetsByPrefix(archive, relsPath, [
      '../diagrams/data',
    ]);
  }
}
