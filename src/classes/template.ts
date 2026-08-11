import { FileHelper } from '../helper/file-helper';
import { ContentTracker } from '../helper/content-tracker';
import { CountHelper } from '../helper/count-helper';
import { ICounter } from '../interfaces/icounter';
import { ISlide } from '../interfaces/islide';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ITemplate } from '../interfaces/itemplate';
import { XmlTemplateHelper } from '../helper/xml-template-helper';
import { ContentMap, SlideInfo, XmlDocument } from '../types/xml-types';
import { XmlHelper } from '../helper/xml-helper';
import { PptPaths } from '../helper/ppt-paths';
import { MediaDeduplicator } from '../helper/media-deduplicator';
import IArchive from '../interfaces/iarchive';
import { ArchiveParams, AutomizerFile, MediaFile } from '../types/types';

import Automizer from '../automizer';
import { IMaster } from '../interfaces/imaster';
import { ILayout } from '../interfaces/ilayout';
import { IGenerator } from '../interfaces/igenerator';

/**
 * Shared base for the two template roles: a source template provides
 * importable slides and shapes (`SourceTemplate`), the root template is the
 * output presentation everything is written into (`OutputTemplate`).
 */
export abstract class Template implements ITemplate {
  /**
   * Path to local file
   * @type string
   */
  location: string;

  /**
   * Node file buffer or path passed to the archive.
   */
  file: AutomizerFile;

  /**
   * this.file will be passed to FileProxy
   * @type Archive
   */
  archive: IArchive;

  protected constructor(file: AutomizerFile, params: ArchiveParams) {
    this.file = file;
    this.archive = FileHelper.importArchive(file, params);
  }

  /**
   * Factory: `params.name` decides the role — a named template is a source
   * template containing importable slides; an unnamed one is the root
   * (output) template.
   */
  static import(
    file: AutomizerFile,
    params: ArchiveParams,
    automizer?: Automizer,
  ): SourceTemplate | OutputTemplate {
    if (params.name) {
      return new SourceTemplate(file, params);
    }
    return new OutputTemplate(file, params, automizer);
  }

  async getSlideIdList(): Promise<XmlDocument> {
    return XmlHelper.getXmlFromArchive(this.archive, PptPaths.presentation);
  }
}

/**
 * A loaded .pptx used as a source of slides, masters and shapes
 * (`Automizer.load()`). Identified by its alias `name`.
 */
export class SourceTemplate extends Template implements PresTemplate {
  /**
   * An alias name to identify template and simplify
   * @type string
   */
  name: string;

  creationIds: SlideInfo[];
  useCreationIds?: boolean;
  slideNumbers: number[];

  constructor(file: AutomizerFile, params: ArchiveParams) {
    super(file, params);
    this.name = params.name;
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
    const xmlTemplateHelper = new XmlTemplateHelper(this.archive);
    this.creationIds = await xmlTemplateHelper.getCreationIds();

    return this.creationIds;
  }
}

/**
 * The root template: the output presentation that slides, masters and media
 * are appended to (`Automizer.loadRoot()`).
 */
export class OutputTemplate extends Template implements RootPresTemplate {
  /**
   * Array containing all slides coming from Automizer.addSlide()
   * @type: ISlide[]
   */
  slides: ISlide[] = [];

  /**
   * Array containing all slideMasters coming from Automizer.addMaster()
   * @type: IMaster[]
   */
  masters: IMaster[] = [];

  /**
   * Array containing all counters
   * @type: ICounter[]
   */
  counter: ICounter[];

  existingSlides: number;

  contentMap: ContentMap[] = [];
  mediaFiles: MediaFile[] = [];

  /**
   * Checksum index of the media files of the output archive, used to import
   * each distinct image only once.
   */
  mediaDeduplicator: MediaDeduplicator;

  content: ContentTracker;
  automizer: Automizer;
  generator: IGenerator;

  constructor(
    file: AutomizerFile,
    params: ArchiveParams,
    automizer?: Automizer,
  ) {
    super(file, params);

    this.automizer = automizer;
    this.counter = [
      new CountHelper('slides', this),
      new CountHelper('charts', this),
      new CountHelper('images', this),
      new CountHelper('diagrams', this),
      new CountHelper('masters', this),
      new CountHelper('layouts', this),
      new CountHelper('themes', this),
      new CountHelper('oleObjects', this),
    ];
    this.content = automizer?.content ?? new ContentTracker();
    this.archive.contentTracker = this.content;
    this.mediaDeduplicator = new MediaDeduplicator(this.archive);
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
      XmlHelper.writeXmlToArchive(this.archive, PptPaths.presentation, xml);

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
      archive: this.archive,
      file: PptPaths.presentationRels,
      tag: 'Relationship',
      clause: (xml, element) =>
        removedRelIds.includes(element.getAttribute('Id')),
    });
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
    const requiresGenerator = this.slides.some(
      (slide) => slide.getGeneratedElements().length > 0,
    );
    if (!requiresGenerator) {
      return;
    }

    // Lazy import: pptxgenjs is a heavy dependency that pure
    // "modify existing pptx" runs never need.
    const { default: GeneratePptxGenJs } = await import(
      '../helper/generate/generate-pptxgenjs'
    );
    this.generator = new GeneratePptxGenJs(this.automizer, this.slides);
    await this.generator.generateSlides();
  }

  async cleanupExternalGenerator() {
    await this.generator?.cleanup();
  }
}
