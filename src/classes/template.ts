import { FileHelper } from '../helper/file-helper';
import { CountHelper } from '../helper/count-helper';
import { ICounter } from '../interfaces/icounter';
import { ISlide } from '../interfaces/islide';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ITemplate } from '../interfaces/itemplate';
import { XmlTemplateHelper } from '../helper/xml-template-helper';
import { SlideInfo } from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { ContentTracker } from '../helper/content-tracker';
import IArchive from '../interfaces/iarchive';
import { ArchiveParams } from '../types/types';
import { IMaster } from '../interfaces/imaster';

export class Template implements ITemplate {
  /**
   * Path to local file
   * @type string
   */
  location: string;

  /**
   * An alias name to identify template and simplify
   * @type string
   */
  name: string;

  /**
   * Node file buffer
   * @type InputType
   */
  file: any;

  /**
   * this.file will be passed to FileProxy
   * @type Archive
   */
  archive: IArchive;

  /**
   * Array containing all slides coming from Automizer.addSlide()
   * @type: ISlide[]
   */
  slides: ISlide[];

  /**
   * Array containing all slides coming from Automizer.addSlide()
   * @type: ISlide[]
   */
  masters: IMaster[];

  /**
   * Array containing all counters
   * @type: ICounter[]
   */
  counter: ICounter[];

  creationIds: SlideInfo[];
  existingSlides: number;
  existingMasterSlides: number;

  constructor(location: string, params: ArchiveParams) {
    this.location = location;
    const archive = FileHelper.importArchive(location, params);
    this.archive = archive;
  }

  static import(
    location: string,
    params: ArchiveParams,
  ): PresTemplate | RootPresTemplate {
    let newTemplate: PresTemplate | RootPresTemplate;
    if (params.name) {
      newTemplate = new Template(location, params) as PresTemplate;
      newTemplate.name = params.name;
    } else {
      newTemplate = new Template(location, params) as RootPresTemplate;
      newTemplate.slides = [];
      newTemplate.masters = [];
      newTemplate.counter = [
        new CountHelper('slides', newTemplate),
        new CountHelper('charts', newTemplate),
        new CountHelper('images', newTemplate),
        new CountHelper('masters', newTemplate),
        new CountHelper('layouts', newTemplate),
      ];
      newTemplate.content = new ContentTracker();
    }

    return newTemplate;
  }

  async setCreationIds(): Promise<SlideInfo[]> {
    const archive = await this.archive;

    const xmlTemplateHelper = new XmlTemplateHelper(archive);
    this.creationIds = await xmlTemplateHelper.getCreationIds();

    return this.creationIds;
  }

  async appendMasterSlide(slideMaster: IMaster): Promise<void> {
    if (this.counter[0].get() === undefined) {
      await this.initializeCounter();
    }

    await slideMaster.append(this);
  }

  async appendSlide(slide: ISlide): Promise<void> {
    if (this.counter[0].get() === undefined) {
      await this.initializeCounter();
    }

    await slide.append(this);
  }

  async countExistingSlides(): Promise<void> {
    const xml = await this.getSlideIdList();
    const sldIdLst = xml.getElementsByTagName('p:sldIdLst');
    if (sldIdLst.length > 0) {
      const existingSlides = sldIdLst[0].getElementsByTagName('p:sldId');
      this.existingSlides = existingSlides.length;
    }
  }

  async truncate(): Promise<void> {
    if (this.existingSlides > 0) {
      const xml = await this.getSlideIdList();
      const existingSlides = xml.getElementsByTagName('p:sldId');
      XmlHelper.sliceCollection(existingSlides, this.existingSlides, 0);
      XmlHelper.writeXmlToArchive(
        await this.archive,
        `ppt/presentation.xml`,
        xml,
      );
    }
  }

  async getSlideIdList(): Promise<Document> {
    const archive = await this.archive;
    const xml = await XmlHelper.getXmlFromArchive(
      archive,
      `ppt/presentation.xml`,
    );
    return xml;
  }

  async initializeCounter(): Promise<void> {
    for (const c of this.counter) {
      await c.set();
    }
  }

  incrementCounter(name: string): number {
    return CountHelper.increment(name, this.counter);
  }

  count(name: string): number {
    return CountHelper.count(name, this.counter);
  }
}
