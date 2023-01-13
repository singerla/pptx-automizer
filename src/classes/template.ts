import JSZip, { InputType } from 'jszip';

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
import { vd } from '../helper/general-helper';
import { ContentTracker } from '../helper/content-tracker';
import { CacheHelper } from '../helper/cache-helper';
import { FileProxy } from '../helper/file-proxy';

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
  file: InputType;

  /**
   * this.file will be passed to FileProxy
   * @type FileProxy
   */
  archive: FileProxy;

  /**
   * Array containing all slides coming from Automizer.addSlide()
   * @type: ISlide[]
   */
  slides: ISlide[];

  /**
   * Array containing all counters
   * @type: ICounter[]
   */
  counter: ICounter[];

  creationIds: SlideInfo[];
  existingSlides: number;

  constructor(location: string) {
    this.location = location;
    const archive = FileHelper.importArchive(location);
    this.archive = archive;
  }

  static import(
    location: string,
    name?: string,
  ): PresTemplate | RootPresTemplate {
    let newTemplate: PresTemplate | RootPresTemplate;
    if (name) {
      newTemplate = new Template(location) as PresTemplate;
      newTemplate.name = name;
    } else {
      newTemplate = new Template(location) as RootPresTemplate;
      newTemplate.slides = [];
      newTemplate.counter = [
        new CountHelper('slides', newTemplate),
        new CountHelper('charts', newTemplate),
        new CountHelper('images', newTemplate),
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
      await XmlHelper.writeXmlToArchive(
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
