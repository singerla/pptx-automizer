import { FileHelper } from '../helper/file-helper';
import { ContentTracker } from '../helper/content-tracker';
import { CountHelper } from '../helper/count-helper';
import { ICounter } from '../interfaces/icounter';
import { ISlide } from '../interfaces/islide';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ITemplate } from '../interfaces/itemplate';
import { XmlTemplateHelper } from '../helper/xml-template-helper';
import { ContentMap, SlideInfo } from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { PptPaths } from '../helper/ppt-paths';
import IArchive from '../interfaces/iarchive';
import { ArchiveParams, AutomizerFile, MediaFile } from '../types/types';

import Automizer from '../automizer';
import { IMaster } from '../interfaces/imaster';
import { ILayout } from '../interfaces/ilayout';
import { IGenerator } from '../interfaces/igenerator';
import GeneratePptxGenJs from '../helper/generate/generate-pptxgenjs';

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
   * Array containing all slideMasters coming from Automizer.addMaster()
   * @type: IMaster[]
   */
  masters: IMaster[];

  /**
   * Array containing all counters
   * @type: ICounter[]
   */
  counter: ICounter[];

  creationIds: SlideInfo[];
  slideNumbers: number[];
  existingSlides: number;

  contentMap: ContentMap[] = [];
  mediaFiles: MediaFile[] = [];

  automizer: Automizer;
  generator: IGenerator;

  constructor(file: AutomizerFile, params: ArchiveParams) {
    this.file = file;
    const archive = FileHelper.importArchive(file, params);
    this.archive = archive;
  }

  static import(
    file: AutomizerFile,
    params: ArchiveParams,
    automizer?: Automizer,
  ): PresTemplate | RootPresTemplate {
    let newTemplate: PresTemplate | RootPresTemplate;
    if (params.name) {
      // New template will be a default template containing
      // importable slides and shapes.
      newTemplate = new Template(file, params) as PresTemplate;
      newTemplate.name = params.name;
    } else {
      // New template will be root template
      newTemplate = new Template(file, params) as RootPresTemplate;
      newTemplate.automizer = automizer;
      newTemplate.slides = [];
      newTemplate.masters = [];
      newTemplate.counter = [
        new CountHelper('slides', newTemplate),
        new CountHelper('charts', newTemplate),
        new CountHelper('images', newTemplate),
        new CountHelper('diagrams', newTemplate),
        new CountHelper('masters', newTemplate),
        new CountHelper('layouts', newTemplate),
        new CountHelper('themes', newTemplate),
        new CountHelper('oleObjects', newTemplate),
      ];
      newTemplate.content = automizer?.content ?? new ContentTracker();
      newTemplate.archive.contentTracker = newTemplate.content;
    }

    return newTemplate;
  }

  mapContents(
    type: 'slideMaster' | 'slideLayout',
    key: string,
    sourceId: number,
    targetId: number,
    name?: string,
  ) {
    this.contentMap.push({
      type,
      key,
      sourceId,
      targetId,
      name,
    });
  }

  getNamedMappedContent(type: 'slideMaster' | 'slideLayout', name: string) {
    return this.contentMap.find(
      (map) => map.type === type && map.name === name,
    );
  }

  getMappedContent(
    type: 'slideMaster' | 'slideLayout',
    key: string,
    sourceId: number,
  ) {
    return this.contentMap.find(
      (map) =>
        map.type === type && map.key === key && map.sourceId === sourceId,
    );
  }

  /**
   * Returns the slide numbers of a given template as a sorted array of integers.
   * @returns {Promise<number[]>} - A promise that resolves to a sorted array of slide numbers in the template.
   */
  async getAllSlideNumbers(): Promise<number[]> {
    try {
      const xmlTemplateHelper = new XmlTemplateHelper(this.archive);
      this.slideNumbers = await xmlTemplateHelper.getAllSlideNumbers();
      return this.slideNumbers;
    } catch (error) {
      throw new Error(error.message);
    }
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

    await slideMaster.append(this).catch((e) => {
      throw e;
    });
  }

  async appendSlide(slide: ISlide): Promise<void> {
    if (this.counter[0].get() === undefined) {
      await this.initializeCounter();
    }

    await slide.append(this).catch((e) => {
      throw e;
    });
  }

  async appendLayout(slideLayout: ILayout): Promise<void> {
    if (this.counter[0].get() === undefined) {
      await this.initializeCounter();
    }

    await slideLayout.append(this).catch((e) => {
      throw e;
    });
  }

  async countExistingSlides(): Promise<void> {
    const xml = await this.getSlideIdList();
    const sldIdLst = xml.getElementsByTagName('p:sldIdLst');
    if (sldIdLst.length > 0) {
      const existingSlides = sldIdLst[0].getElementsByTagName('p:sldId');
      this.existingSlides = existingSlides.length;
    }
  }

  /**
   * Remove the slides that came with the root template from the presentation,
   * keeping the slides added by automizer. Used by `removeExistingSlides`.
   *
   * Along with the `p:sldId` entries, the corresponding relationships in
   * ppt/_rels/presentation.xml.rels are dropped: a slide part that is still
   * related to the presentation counts as a slide for anything reading the
   * output (including automizer's own `getInfo()`), even if it is not listed
   * in `p:sldIdLst` (see #166). The slide parts themselves are removed by
   * `ModifyPresentationHelper.removeUnusedFiles` if `cleanup` is enabled.
   */
  async truncate(): Promise<void> {
    if (this.existingSlides > 0) {
      const xml = await this.getSlideIdList();
      const existingSlides = xml.getElementsByTagName('p:sldId');

      const removedRelIds: string[] = [];
      const removeCount = Math.min(this.existingSlides, existingSlides.length);
      for (let i = 0; i < removeCount; i++) {
        removedRelIds.push(existingSlides[i].getAttribute('r:id'));
      }

      XmlHelper.sliceCollection(existingSlides, this.existingSlides, 0);
      XmlHelper.writeXmlToArchive(
        await this.archive,
        PptPaths.presentation,
        xml,
      );

      await this.removeSlideRelations(removedRelIds);

      // The slides counter was initialized before appending and still
      // includes the removed slides.
      CountHelper.decrement('slides', this.counter, removeCount);
    }
  }

  /**
   * Remove the given relationship ids from ppt/_rels/presentation.xml.rels.
   */
  async removeSlideRelations(removedRelIds: string[]): Promise<void> {
    if (!removedRelIds.length) {
      return;
    }

    await XmlHelper.removeIf({
      archive: await this.archive,
      file: PptPaths.presentationRels,
      tag: 'Relationship',
      clause: (xml, element) =>
        removedRelIds.includes(element.getAttribute('Id')),
    });
  }

  async getSlideIdList(): Promise<Document> {
    const archive = await this.archive;
    const xml = await XmlHelper.getXmlFromArchive(
      archive,
      PptPaths.presentation,
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

  async runExternalGenerator() {
    this.generator = new GeneratePptxGenJs(this.automizer, this.slides);
    await this.generator.generateSlides();
  }

  async cleanupExternalGenerator() {
    await this.generator.cleanup();
  }
}
