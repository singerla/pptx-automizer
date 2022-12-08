import JSZip from 'jszip';

import { XmlHelper } from '../helper/xml-helper';
import { GeneralHelper } from '../helper/general-helper';
import {
  ImportedElement,
  ShapeModificationCallback,
  Workbook,
} from '../types/types';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { HelperElement } from '../types/xml-types';
import { ImageTypeMap } from '../enums/image-type-map';
import { ElementSubtype } from '../enums/element-type';

export class Shape {
  mode: string;
  name: string;

  sourceArchive: JSZip;
  sourceSlideNumber: number;
  sourceSlideFile: string;
  sourceNumber: number;
  sourceFile: string;
  sourceRid: string;
  sourceElement: XMLDocument;

  targetFile: string;
  targetArchive: JSZip;
  targetTemplate: RootPresTemplate;
  targetSlideNumber: number;
  targetNumber: number;
  targetSlideFile: string;
  targetSlideRelFile: string;

  createdRid: string;

  relRootTag: string;
  relAttribute: string;
  relParent: (element: Element) => Element;

  targetElement: XMLDocument;

  callbacks: ShapeModificationCallback[];
  hasCreationId: boolean;
  contentTypeMap: typeof ImageTypeMap;
  subtype: ElementSubtype;

  constructor(shape: ImportedElement) {
    this.mode = shape.mode;
    this.name = shape.name;

    this.sourceArchive = shape.sourceArchive;
    this.sourceSlideNumber = shape.sourceSlideNumber;
    this.sourceSlideFile = `ppt/slides/slide${this.sourceSlideNumber}.xml`;
    this.sourceElement = shape.sourceElement;
    this.hasCreationId = shape.hasCreationId;

    this.callbacks = GeneralHelper.arrayify(shape.callback);
    this.contentTypeMap = ImageTypeMap;

    if (shape.target) {
      this.sourceNumber = shape.target.number;
      this.sourceRid = shape.target.rId;
      this.subtype = shape.target.subtype;
    }
  }

  async setTarget(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    this.targetTemplate = targetTemplate;
    this.targetArchive = await this.targetTemplate.archive;
    this.targetSlideNumber = targetSlideNumber;
    this.targetSlideFile = `ppt/slides/slide${this.targetSlideNumber}.xml`;
    this.targetSlideRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`;
  }

  async setTargetElement(): Promise<void> {
    this.targetElement = this.sourceElement.cloneNode(true) as XMLDocument;
  }

  async appendToSlideTree(): Promise<void> {
    const targetSlideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideFile,
    );

    targetSlideXml
      .getElementsByTagName('p:spTree')[0]
      .appendChild(this.targetElement);

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      this.targetSlideFile,
      targetSlideXml,
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

    const sourceElementOnTargetSlide = await XmlHelper[findMethod](
      targetSlideXml,
      this.name,
    );

    if (insertBefore === true) {
      sourceElementOnTargetSlide.parentNode.insertBefore(
        this.targetElement,
        sourceElementOnTargetSlide,
      );
    }

    sourceElementOnTargetSlide.parentNode.removeChild(
      sourceElementOnTargetSlide,
    );

    // vd(insertBefore);

    await XmlHelper.writeXmlToArchive(archive, slideFile, targetSlideXml);
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

    targetElements.forEach((targetElement) => {
      this.relParent(targetElement)
        .getElementsByTagName(this.relRootTag)[0]
        .setAttribute(this.relAttribute, this.createdRid);
    });

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      this.targetSlideFile,
      targetSlideXml,
    );
  }

  async getElementsByRid(slideXml: Document, rId: string): Promise<Element[]> {
    const sourceList = slideXml
      .getElementsByTagName('p:spTree')[0]
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
      .getElementsByTagName(this.relRootTag)[0]
      .setAttribute(this.relAttribute, this.createdRid);
  }

  applyCallbacks(
    callbacks: ShapeModificationCallback[],
    element: XMLDocument,
    arg1?: Document,
    arg2?: Workbook,
  ): void {
    callbacks.forEach((callback) => {
      if (typeof callback === 'function') {
        try {
          callback(element, arg1, arg2);
        } catch (e) {
          console.log(e);
        }
      }
    });
  }

  appendImageExtensionToContentType(
    extension,
  ): Promise<HelperElement | boolean> {
    const contentType = this.contentTypeMap[extension]
      ? this.contentTypeMap[extension]
      : 'image/' + extension;

    return XmlHelper.appendIf({
      ...XmlHelper.createContentTypeChild(this.targetArchive, {
        Extension: extension,
        ContentType: contentType,
      }),
      tag: 'Default',
      clause: (xml: XMLDocument) =>
        !XmlHelper.findByAttribute(xml, 'Default', 'Extension', extension),
    });
  }
}
