import { Slide } from './classes/slide';
import {
  ArchiveParams,
  AutomizerFile,
  AutomizerParams,
  AutomizerSummary,
  PresentationInfo,
  SourceIdentifier,
  StatusTracker,
} from './types/types';
import { IPresentationProps } from './interfaces/ipresentation-props';
import { PresTemplate } from './interfaces/pres-template';
import { RootPresTemplate } from './interfaces/root-pres-template';
import { Template } from './classes/template';
import { ModifyXmlCallback, TemplateInfo } from './types/xml-types';
import { GeneralHelper, log, Logger } from './helper/general-helper';
import { Master } from './classes/master';
import path from 'path';
import * as fs from 'fs';
import { XmlHelper } from './helper/xml-helper';
import ModifyPresentationHelper from './helper/modify-presentation-helper';
import { contentTracker as Tracker, ContentTracker } from './helper/content-tracker';
import JSZip from 'jszip';
import { ISlide } from './interfaces/islide';
import { IMaster } from './interfaces/imaster';
import { ContentTypeExtension } from './enums/content-type-map';
import slugify from 'slugify';

/**
 * Automizer
 *
 * The basic class for `pptx-automizer` package.
 * This class will be exported as `Automizer` by `index.ts`.
 */
export default class Automizer implements IPresentationProps {
  rootTemplate: RootPresTemplate;
  /**
   * Templates  of automizer
   * @internal
   */
  templates: PresTemplate[] = [];
  templateDir: string;
  templateFallbackDir: string;
  outputDir: string;
  archiveParams: ArchiveParams;
  /**
   * Timer of automizer
   * @internal
   */
  timer: number;
  params: AutomizerParams;
  status: StatusTracker;

  content: ContentTracker;
  modifyPresentation: ModifyXmlCallback[] = [];

  /**
   * Creates an instance of `pptx-automizer`.
   * @param [params]
   */
  constructor(params: AutomizerParams) {
    this.params = params;

    this.templateDir = params?.templateDir ? params.templateDir + '/' : '';
    this.templateFallbackDir = params?.templateFallbackDir
      ? params.templateFallbackDir + '/'
      : '';
    this.outputDir = params?.outputDir ? params.outputDir + '/' : '';

    this.archiveParams = <ArchiveParams>{
      mode: params?.archiveType?.mode || 'jszip',
      baseDir: params?.archiveType?.baseDir || __dirname + '/../cache',
      workDir: params?.archiveType?.workDir || 'tmp',
      cleanupWorkDir: params?.archiveType?.cleanupWorkDir,
    };

    this.timer = Date.now();
    this.setStatusTracker(params?.statusTracker);

    this.content = new ContentTracker();

    if (params.rootTemplate) {
      let file = params.rootTemplate;
      if (typeof file !== 'object') {
        file = this.getLocation(file, 'template');
      }
      this.rootTemplate = Template.import(
        file,
        this.archiveParams,
        this,
      ) as RootPresTemplate;
    }

    if (params.presTemplates) {
      this.params.presTemplates.forEach((file, i) => {
        let name: string;
        if (typeof file !== 'object') {
          name = file;
          file = this.getLocation(file, 'template');
        } else {
          name = `${i}.pptx`;
        }
        const archiveParams = {
          ...this.archiveParams,
          name,
        };
        const newTemplate = Template.import(
          file,
          archiveParams,
        ) as PresTemplate;

        this.templates.push(newTemplate);
      });
    }

    if (params.verbosity) {
      Logger.verbosity = params.verbosity;
    }
  }

  setStatusTracker(statusTracker: StatusTracker['next']): void {
    const defaultStatusTracker = (status: StatusTracker) => {
      log(status.info + ' (' + status.share + '%)', 2);
    };

    this.status = {
      current: 0,
      max: 0,
      share: 0,
      info: undefined,
      increment: () => {
        this.status.current++;
        const nextShare =
          this.status.max > 0
            ? Math.round((this.status.current / this.status.max) * 100)
            : 0;

        if (this.status.share !== nextShare) {
          this.status.share = nextShare;
          this.status.next(this.status);
        }
      },
      next: statusTracker || defaultStatusTracker,
    };
  }

  /**

   */
  public async presentation(): Promise<this> {
    if (this.params?.useCreationIds === true) {
      await this.setCreationIds();
    }
    return this;
  }

  /**
   * Load a pptx file and set it as root template.
   * @param file - Filename, path to the template or Buffer containing the file.
   * Filenames and paths will be prefixed with 'templateDir'
   * @returns Instance of Automizer
   */
  public loadRoot(file: AutomizerFile): this {
    return this.loadTemplate(file);
  }

  /**
   * Load a template pptx file.
   * @param file - Filename, path to the template or Buffer containing the file.
   * Filenames and paths will be prefixed with 'templateDir'
   * @param name - Optional short name for a template loaded from a file. If skipped, the template will be named by its location.
   * if the file is a Buffer the name is required.
   * @returns Instance of Automizer
   */
  public load(file: AutomizerFile, name?: string): this {
    if (!name && typeof file !== 'object') {
      name = name === undefined ? file : name;
    } else if (typeof file === 'object' && !name) {
      throw new Error('Name is required when loading a template from a Buffer');
    }
    return this.loadTemplate(file, name);
  }

  /**
   * Loads a pptx file either as a root template as a template file.
   * A name can be specified to give templates an alias.
   * @param location
   * @param [name]
   * @returns template
   */
  private loadTemplate(file: AutomizerFile, name?: string): this {
    if (typeof file !== 'object') {
      file = this.getLocation(file, 'template');
    }
    const alreadyLoaded = this.templates.find(
      (template) => template.name === name,
    );
    if (alreadyLoaded) {
      return this;
    }

    const importParams = {
      ...this.archiveParams,
      name,
    };

    const newTemplate = Template.import(file, importParams, this);

    if (!this.isPresTemplate(newTemplate)) {
      this.rootTemplate = newTemplate;
    } else {
      this.templates.push(newTemplate);
    }

    return this;
  }

  /**
   * Load media files to output presentation.
   * @returns Instance of Automizer
   * @param filename Filename or path to the media file.
   * @param dir Specify custom path for media instead of mediaDir from AutomizerParams.
   */
  public loadMedia(
    filename: string | string[],
    dir?: string,
    prefix?: string,
  ): this {
    const files = GeneralHelper.arrayify(filename);
    if (!this.rootTemplate) {
      throw "Can't load media, you need to load a root template first";
    }
    files.forEach((file) => {
      const directory = dir || this.params.mediaDir;
      const filepath = path.join(directory, file);
      const extension = path
        .extname(file)
        .replace('.', '') as ContentTypeExtension;
      try {
        fs.accessSync(filepath, fs.constants.F_OK);
      } catch (e) {
        throw `Can't load media: ${filepath} does not exist.`;
      }
      this.rootTemplate.mediaFiles.push({
        file,
        directory,
        filepath,
        extension,
        prefix,
      });
    });
    return this;
  }

  /**
   * Parses all loaded templates and collects creationIds for slides and
   * elements. This will make finding templates and elements independent
   * of slide number and element name.
   * @returns Promise<TemplateInfo[]>
   */
  public async setCreationIds(): Promise<TemplateInfo[]> {
    const templateCreationId = [];
    for (const template of this.templates) {
      const creationIds =
        template.creationIds || (await template.setCreationIds());
      template.useCreationIds = this.params.useCreationIds;
      templateCreationId.push({
        name: template.name,
        slides: creationIds,
      });
    }
    return templateCreationId;
  }

  /**
   * Get some info about the imported templates
   * @returns Promise<PresentationInfo>
   */
  public async getInfo(): Promise<PresentationInfo> {
    const creationIds = await this.setCreationIds();
    const info: PresentationInfo = {
      templateByName: (tplName: string) => {
        return creationIds.find((template) => template.name === tplName);
      },
      slidesByTemplate: (tplName: string) => {
        return info.templateByName(tplName)?.slides || [];
      },
      slideByNumber: (tplName: string, slideNumber: number) => {
        return info
          .slidesByTemplate(tplName)
          .find((slide) => slide.number === slideNumber);
      },
      elementByName: (
        tplName: string,
        slideNumber: number,
        elementName: string,
      ) => {
        return info
          .slideByNumber(tplName, slideNumber)
          ?.elements.find((element) => elementName === element.name);
      },
    };
    return info;
  }

  /**
   * Determines whether template is root or default template.
   * @param template
   * @returns pres template
   */
  private isPresTemplate(
    template: PresTemplate | RootPresTemplate,
  ): template is PresTemplate {
    return 'name' in template;
  }

  /**
   * Add a slide from one of the imported templates by slide number or creationId.
   * @param name - Name or alias of the template; must have been loaded with `Automizer.load()`
   * @param slideIdentifier - Number or creationId of slide in template presentation
   * @param callback - Executed after slide was added. The newly created slide will be passed to the callback as first argument.
   * @returns Instance of Automizer
   */
  public addSlide(
    name: string,
    slideIdentifier: SourceIdentifier,
    callback?: (slide: ISlide) => void,
  ): this {
    if (this.rootTemplate === undefined) {
      throw new Error('You have to set a root template first.');
    }

    const template = this.getTemplate(name);

    const newSlide = new Slide({
      presentation: this,
      template,
      slideIdentifier,
    });

    if (this.params.autoImportSlideMasters) {
      newSlide.useSlideLayout();
    }

    if (callback !== undefined) {
      newSlide.root = this;
      callback(newSlide);
    }

    this.rootTemplate.slides.push(newSlide);

    return this;
  }

  /**
   * Copy and modify a master and the associated layouts from template to output.
   *
   * @param name
   * @param sourceIdentifier
   * @param callback
   */
  public addMaster(
    name: string,
    sourceIdentifier: number,
    callback?: (slideMaster: IMaster) => void,
  ): this {
    const key = sourceIdentifier + '@' + name;

    if (this.rootTemplate.masters.find((master) => master.key === key)) {
      console.log('Already imported ' + key);
      return this;
    }

    const template = this.getTemplate(name);

    const newMaster = new Master({
      presentation: this,
      template,
      sourceIdentifier,
    });

    if (callback !== undefined) {
      newMaster.root = this;
      callback(newMaster);
    }

    this.rootTemplate.masters.push(newMaster);

    return this;
  }

  /**
   * Searches this.templates to find template by given name.
   * @internal
   * @param name Alias name if given to loaded template.
   * @returns template
   */
  public getTemplate(name: string): PresTemplate {
    const template = this.templates.find((t) => t.name === name);
    if (template === undefined) {
      throw new Error(`Template not found: ${name}`);
    }
    return template;
  }

  /**
   * Write all imports and modifications to a file.
   * @param location - Filename or path for the file. Will be prefixed with 'outputDir'
   * @returns summary object.
   */
  public async write(location: string): Promise<AutomizerSummary> {
    await this.finalizePresentation();

    await this.rootTemplate.archive.output(
      this.getLocation(location, 'output'),
      this.params,
    );

    const duration: number = (Date.now() - this.timer) / 600;

    return {
      status: 'finished',
      duration,
      file: location,
      filename: path.basename(location),
      templates: this.templates.length,
      slides: this.rootTemplate.count('slides'),
      charts: this.rootTemplate.count('charts'),
      images: this.rootTemplate.count('images'),
      masters: this.rootTemplate.count('masters'),
    };
  }

  /**
   * Create a ReadableStream from output pptx file.
   * @param generatorOptions - JSZipGeneratorOptions for nodebuffer Output type
   * @returns Promise<NodeJS.ReadableStream>
   */
  public async stream(
    generatorOptions?: JSZip.JSZipGeneratorOptions<'nodebuffer'>,
  ): Promise<NodeJS.ReadableStream> {
    await this.finalizePresentation();

    if (!this.rootTemplate.archive.stream) {
      throw 'Streaming is not implemented for current archive type';
    }

    return this.rootTemplate.archive.stream(this.params, generatorOptions);
  }

  /**
   * Pass final JSZip instance.
   * @returns Promise<NodeJS.ReadableStream>
   */
  public async getJSZip(): Promise<JSZip> {
    await this.finalizePresentation();

    if (!this.rootTemplate.archive.getFinalArchive) {
      throw 'GetFinalArchive is not implemented for current archive type';
    }

    return this.rootTemplate.archive.getFinalArchive();
  }

  async finalizePresentation() {
    await this.writeMasterSlides();
    await this.writeSlides();
    await this.writeMediaFiles();
    await this.normalizePresentation();
    await this.applyModifyPresentationCallbacks();

    // TODO: refactor content tracker, move this to root template
    Tracker.reset();
  }

  /**
   * Write all masterSlides to archive.
   */
  public async writeMasterSlides(): Promise<void> {
    for (const slide of this.rootTemplate.masters) {
      await this.rootTemplate.appendMasterSlide(slide);
    }
  }

  /**
   * Write all slides to archive.
   */
  public async writeSlides(): Promise<void> {
    await this.rootTemplate.countExistingSlides();
    this.status.max = this.rootTemplate.slides.length;

    await this.rootTemplate.runExternalGenerators();
    for (const slide of this.rootTemplate.slides) {
      await this.rootTemplate.appendSlide(slide);
    }
    await this.rootTemplate.cleanupExternalGenerators();

    if (this.params.removeExistingSlides) {
      await this.rootTemplate.truncate();
    }
  }

  /**
   * Write all media files to archive.
   */
  public async writeMediaFiles(): Promise<void> {
    const mediaDir = 'ppt/media/';
    for (const file of this.rootTemplate.mediaFiles) {
      const data = fs.readFileSync(file.filepath);
      let archiveFilename = file.file;
      if (file.prefix) {
        archiveFilename = file.prefix + file.file;
      }

      archiveFilename = slugify(archiveFilename);

      await this.rootTemplate.archive.write(mediaDir + archiveFilename, data);
      await XmlHelper.appendImageExtensionToContentType(
        this.rootTemplate.archive,
        file.extension,
      );
    }
  }

  /**
   * Applies all callbacks in this.modifyPresentation-array.
   * The callback array can be pushed by this.modify()
   */
  async applyModifyPresentationCallbacks(): Promise<void> {
    await XmlHelper.modifyXmlInArchive(
      this.rootTemplate.archive,
      `ppt/presentation.xml`,
      this.modifyPresentation,
    );
  }

  /**
   * Apply some callbacks to restore archive/xml structure
   * and prevent corrupted pptx files.
   *
   * TODO: Use every imported image only once
   * TODO: Check for lost relations
   */
  async normalizePresentation(): Promise<void> {
    this.modify(ModifyPresentationHelper.normalizeSlideIds);
    this.modify(ModifyPresentationHelper.normalizeSlideMasterIds);

    if (this.params.cleanup) {
      if (this.params.removeExistingSlides) {
        this.modify(ModifyPresentationHelper.removeUnusedFiles);
      }
      this.modify(ModifyPresentationHelper.removedUnusedImages);
      this.modify(ModifyPresentationHelper.removeUnusedContentTypes);
    }
  }

  public modify(cb: ModifyXmlCallback): this {
    this.modifyPresentation.push(cb);
    return this;
  }

  /**
   * Applies path prefix to given location string.
   * @param location path and/or filename
   * @param [type] template or output
   * @returns location
   */
  private getLocation(location: string, type?: string): string {
    switch (type) {
      case 'template':
        if (fs.existsSync(this.templateDir + location)) {
          return this.templateDir + location;
        } else if (fs.existsSync(this.templateFallbackDir + location)) {
          return this.templateFallbackDir + location;
        } else {
          if (typeof location === 'string') {
            log('No file matches "' + location + '"', 0);
          } else {
            log('Invalid filename', 0);
          }

          log('@templateDir: ' + this.templateDir, 2);
          log('@templateFallbackDir: ' + this.templateFallbackDir, 2);
        }
        break;
      case 'output':
        return this.outputDir + location;
      default:
        return location;
    }
  }
}
