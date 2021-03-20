import JSZip from 'jszip';
import XmlHelper from './helper/xml';
import { ImportedElement, RootPresTemplate, Target } from './definitions/app';
import GeneralHelper from './helper/general';

export default class Shape {
  mode: string;
  name: string;

  sourceArchive: JSZip;
  sourceSlideNumber: number;
  sourceSlideFile: string;
  sourceNumber: number;
  sourceFile: string;
  sourceRid: string;
  sourceElement: HTMLElement;

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
  relParent: (element: HTMLElement) => HTMLElement;

  targetElement: HTMLElement;
  callbacks: Function[];

  constructor(shape: ImportedElement) {
    this.mode = shape.mode;
    this.name = shape.name;

    this.sourceArchive = shape.sourceArchive;
    this.sourceSlideNumber = shape.sourceSlideNumber;
    this.sourceSlideFile = `ppt/slides/slide${this.sourceSlideNumber}.xml`;
    this.sourceElement = shape.sourceElement;

    this.callbacks = GeneralHelper.arrayify(shape.callback);

    if (shape.target) {
      this.sourceNumber = shape.target.number;
      this.sourceRid = shape.target.rId;
    }
  }

  async setTarget(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<void> {
    this.targetTemplate = targetTemplate;
    this.targetArchive = await this.targetTemplate.archive;
    this.targetSlideNumber = targetSlideNumber;
    this.targetSlideFile = `ppt/slides/slide${this.targetSlideNumber}.xml`;
    this.targetSlideRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`;
  }

  async setTargetElement(): Promise<void> {
    this.targetElement = <HTMLElement>this.sourceElement.cloneNode(true);
  }

  async appendToSlideTree(): Promise<void> {
    let targetSlideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile);
    targetSlideXml.getElementsByTagName('p:spTree')[0].appendChild(this.targetElement);

    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml);
  }

  async replaceIntoSlideTree(): Promise<void> {
    let targetSlideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile);

    let sourceElementOnTargetSlide = await XmlHelper.findByName(targetSlideXml, this.name);

    sourceElementOnTargetSlide.parentNode.insertBefore(this.targetElement, sourceElementOnTargetSlide);
    sourceElementOnTargetSlide.parentNode.removeChild(sourceElementOnTargetSlide);

    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml);
  }

  async updateElementRelId() {
    let targetSlideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile);
    let targetElement = await this.getElementByRid(targetSlideXml, this.sourceRid);
    targetElement.getElementsByTagName(this.relRootTag)[0].setAttribute(this.relAttribute, this.createdRid);


    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml);
  }

  async updateTargetElementRelId() {
    this.targetElement.getElementsByTagName(this.relRootTag)[0].setAttribute(this.relAttribute, this.createdRid);
  }

  async getElementByRid(slideXml: Document, rId: string): Promise<HTMLElement> {
    let sourceList = slideXml.getElementsByTagName('p:spTree')[0].getElementsByTagName(this.relRootTag);
    let sourceElement = XmlHelper.findByAttributeValue(sourceList, this.relAttribute, rId);

    return this.relParent(sourceElement);
  }

  applyCallbacks(callbacks: Function[], element: HTMLElement, arg1?: any, arg2?: any): void {
    for (let i in callbacks) {
      callbacks[i](element, arg1, arg2);
    }
  }
}
