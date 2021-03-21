import JSZip from 'jszip';
import Chart from './shapes/chart';
import Image from './shapes/image';

import FileHelper from './helper/file';
import XmlHelper from './helper/xml';

import { ElementType } from './definitions/enums';
import {
  ISlide,
  RootPresTemplate,
  PresTemplate,
  IPresentationProps,
  ImportedElement,
  AnalyzedElementType,
  Target,
  ImportElement
} from './definitions/app';
import { RelationshipAttribute, SlideListAttribute } from './definitions/xml';
import Generic from './shapes/generic';
import { ModificationCallback } from './definitions/types';

export default class Slide implements ISlide {
  sourceTemplate: PresTemplate;
  targetTemplate: RootPresTemplate;
  targetNumber: number;
  sourceNumber: number;
  targetArchive: JSZip;
  sourceArchive: JSZip;
  sourcePath: string;
  targetPath: string;
  modifications: ModificationCallback[];
  importElements: ImportElement[];
  relsPath: string;
  rootTemplate: RootPresTemplate;
  root: IPresentationProps;
  targetRelsPath: string;

  constructor(params: any) {
    this.sourceTemplate = params.template;
    this.sourceNumber = params.number;
    this.sourcePath = `ppt/slides/slide${this.sourceNumber}.xml`;
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`;

    this.modifications = [];
    this.importElements = [];
  }

  async append(targetTemplate: RootPresTemplate) {
    this.targetTemplate = targetTemplate;
    this.targetArchive = await targetTemplate.archive;
    this.targetNumber = targetTemplate.incrementCounter('slides');
    this.targetPath = `ppt/slides/slide${this.targetNumber}.xml`;
    this.targetRelsPath = `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`;

    this.sourceArchive = await this.sourceTemplate.archive;

    await this.copySlideFiles();
    await this.copyRelatedContent();
    await this.addSlideToPresentation();

    if (this.hasNotes()) {
      await this.copySlideNoteFiles();
      await this.updateSlideNoteFile();
      await this.appendNotesToContentType(this.targetArchive, this.targetNumber);
    }

    if (this.importElements.length) {
      await this.importedSelectedElements();
    }

    await this.applyModifications();
  }

  modify(callback: ModificationCallback): void {
    this.modifications.push(callback);
  }

  async addSlideToPresentation(): Promise<void> {
    const relId = await XmlHelper.getNextRelId(this.targetArchive, 'ppt/_rels/presentation.xml.rels');
    await this.appendToSlideRel(this.targetArchive, relId, this.targetNumber),
      await this.appendToSlideList(this.targetArchive, relId),
      await this.appendSlideToContentType(this.targetArchive, this.targetNumber);
  }

  async modifyElement(selector: string, callback: Function | Function[]): Promise<this> {
    const presName = this.sourceTemplate.name;
    const slideNumber = this.sourceNumber;

    this.addElementToModlist(presName, slideNumber, selector, 'modify', callback);

    return this;
  }

  async addElement(presName: string, slideNumber: number, selector: string, callback?: Function | Function[]): Promise<this> {
    this.addElementToModlist(presName, slideNumber, selector, 'append', callback);

    return this;
  }

  async addElementToModlist(presName: string, slideNumber: number, selector: string, mode: string, callback?: Function | Function[]): Promise<this> {
    this.importElements.push({
      presName,
      slideNumber,
      selector,
      mode,
      callback
    });

    return this;
  }

  async importedSelectedElements(): Promise<void> {
    for (const element of this.importElements) {
      const info = await this.getElementInfo(element);

      switch (info.type) {
        case ElementType.Chart:
          await new Chart(info)[info.mode](this.targetTemplate, this.targetNumber);
          break;
        case ElementType.Image:
          await new Image(info)[info.mode](this.targetTemplate, this.targetNumber);
          break;
        case ElementType.Shape:
          await new Generic(info)[info.mode](this.targetTemplate, this.targetNumber);
          break;
      }
    }
  }

  async getElementInfo(importElement: ImportElement): Promise<ImportedElement> {
    const template = this.root.template(importElement.presName);
    const sourcePath = `ppt/slides/slide${importElement.slideNumber}.xml`;
    const sourceArchive = await template.archive;
    const sourceElement = await XmlHelper.findByElementName(sourceArchive, sourcePath, importElement.selector);

    if (!sourceElement) {
      throw new Error(`Can't find ${importElement.selector} on slide ${importElement.slideNumber} in ${importElement.presName}`);
    }

    const appendElementParams = await this.analyzeElement(sourceElement, sourceArchive, importElement.slideNumber);

    return {
      mode: importElement.mode,
      name: importElement.selector,
      sourceArchive,
      sourceSlideNumber: importElement.slideNumber,
      sourceElement,
      callback: importElement.callback,
      target: appendElementParams.target,
      type: appendElementParams.type,
    };
  }

  async analyzeElement(sourceElement: any, sourceArchive: JSZip, slideNumber: number): Promise<AnalyzedElementType> {
    const isChart = sourceElement.getElementsByTagName('c:chart');
    if (isChart.length) {
      return {
        type: ElementType.Chart,
        target: await XmlHelper.getTargetByRelId(sourceArchive, slideNumber, sourceElement, 'chart'),
      } as AnalyzedElementType;
    }

    const isImage = sourceElement.getElementsByTagName('p:nvPicPr');
    if (isImage.length) {
      return {
        type: ElementType.Image,
        target: await XmlHelper.getTargetByRelId(sourceArchive, slideNumber, sourceElement, 'image'),
      } as AnalyzedElementType;
    }

    return {
      type: ElementType.Shape,
    } as AnalyzedElementType;
  }

  async applyModifications(): Promise<void> {
    for (const modification of this.modifications) {
      const xml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetPath);
      modification(xml);
      await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, xml);
    }
  }

  async copySlideFiles(): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/slides/slide${this.sourceNumber}.xml`,
      this.targetArchive, `ppt/slides/slide${this.targetNumber}.xml`
    );

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`,
      this.targetArchive, `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`
    );
  }

  async copySlideNoteFiles(): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/notesSlides/notesSlide${this.sourceNumber}.xml`,
      this.targetArchive, `ppt/notesSlides/notesSlide${this.targetNumber}.xml`
    );

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/notesSlides/_rels/notesSlide${this.sourceNumber}.xml.rels`,
      this.targetArchive, `ppt/notesSlides/_rels/notesSlide${this.targetNumber}.xml.rels`
    );
  }

  async updateSlideNoteFile(): Promise<void> {
    await XmlHelper.replaceAttribute(
      this.targetArchive,
      `ppt/notesSlides/_rels/notesSlide${this.targetNumber}.xml.rels`, 'Relationship', 'Target',
      `../slides/slide${this.sourceNumber}.xml`,
      `../slides/slide${this.targetNumber}.xml`
    );

    await XmlHelper.replaceAttribute(
      this.targetArchive,
      `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`, 'Relationship', 'Target',
      `../notesSlides/notesSlide${this.sourceNumber}.xml`,
      `../notesSlides/notesSlide${this.targetNumber}.xml`
    );
  }

  appendToSlideRel(rootArchive: JSZip, relId: string, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/_rels/presentation.xml.rels`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`,
        Target: `slides/slide${slideCount}.xml`
      } as RelationshipAttribute
    });
  }

  appendToSlideList(rootArchive: JSZip, relId: string): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/presentation.xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('p:sldIdLst')[0],
      tag: 'p:sldId',
      attributes: {
        id: (xml: HTMLElement) => XmlHelper.getMaxId(xml.getElementsByTagName('p:sldId'), 'id', true),
        'r:id': relId
      } as SlideListAttribute
    });
  }

  appendSlideToContentType(rootArchive: JSZip, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/slides/slide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`
      })
    );
  }

  appendNotesToContentType(rootArchive: JSZip, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/notesSlides/notesSlide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml`
      })
    );
  }

  async copyRelatedContent(): Promise<void> {
    const charts = await Chart.getAllOnSlide(this.sourceArchive, this.relsPath);
    for (const chart of charts) {
      await new Chart({
        mode: 'append',
        target: chart,
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceNumber,
      }).modifyOnAddedSlide(this.targetTemplate, this.targetNumber);
    }

    const images = await Image.getAllOnSlide(this.sourceArchive, this.relsPath);
    for (const image of images) {
      await new Image({
        mode: 'append',
        target: image,
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceNumber,
      }).modifyOnAddedSlide(this.targetTemplate, this.targetNumber);
    }
  }

  hasNotes(): boolean {
    const file = this.sourceArchive.file(`ppt/notesSlides/notesSlide${this.sourceNumber}.xml`);
    return file && file.hasOwnProperty('name');
  }

}
