import { Slide } from './classes/slide';
import { FileHelper } from './helper/file-helper';
import {
  AutomizerParams,
  AutomizerSummary,
  SourceSlideIdentifier,
  StatusTracker,
} from './types/types';
import { IPresentationProps } from './interfaces/ipresentation-props';
import { PresTemplate } from './interfaces/pres-template';
import { RootPresTemplate } from './interfaces/root-pres-template';
import { Template } from './classes/template';
import { ModifyXmlCallback, TemplateInfo } from './types/xml-types';
import { vd } from './helper/general-helper';
import { Master } from './classes/master';
import path from 'path';
import * as fs from 'fs';
import { XmlHelper } from './helper/xml-helper';
import ModifyPresentationHelper from './helper/modify-presentation-helper';
import { contentTracker, ContentTracker } from './helper/content-tracker';
import JSZip from 'jszip';

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
  templates: PresTemplate[];
  templateDir: string;
  templateFallbackDir: string;
  outputDir: string;
  /**
   * Timer  of automizer
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
    this.templates = [];
    this.params = params;

    this.templateDir = params?.templateDir ? params.templateDir + '/' : '';
    this.templateFallbackDir = params?.templateFallbackDir
      ? params.templateFallbackDir + '/'
      : '';
    this.outputDir = params?.outputDir ? params.outputDir + '/' : '';

    this.timer = Date.now();
    this.setStatusTracker(params?.statusTracker);

    this.content = new ContentTracker();

    if (params.rootTemplate) {
      const location = this.getLocation(params.rootTemplate, 'template');
      this.rootTemplate = Template.import(location) as RootPresTemplate;
    }

    if (params.presTemplates) {
      this.params.presTemplates.forEach((file) => {
        const location = this.getLocation(file, 'template');
        const newTemplate = Template.import(location, file) as PresTemplate;
        this.templates.push(newTemplate);
      });
    }
  }

  setStatusTracker(statusTracker: StatusTracker['next']): void {
    const defaultStatusTracker = (status: StatusTracker) => {
      console.log(status.info + ' (' + status.share + '%)');
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
   * @param location - Filename or path to the template. Will be prefixed with 'templateDir'
   * @returns Instance of Automizer
   */
  public loadRoot(location: string): this {
    return this.loadTemplate(location);
  }

  /**
   * Load a template pptx file.
   * @param location - Filename or path to the template. Will be prefixed with 'templateDir'
   * @param name - Optional: A short name for the template. If skipped, the template will be named by its location.
   * @returns Instance of Automizer
   */
  public load(location: string, name?: string): this {
    name = name === undefined ? location : name;
    return this.loadTemplate(location, name);
  }

  /**
   * Loads a pptx file either as a root template as a template file.
   * A name can be specified to give templates an alias.
   * @param location
   * @param [name]
   * @returns template
   */
  private loadTemplate(location: string, name?: string): this {
    location = this.getLocation(location, 'template');
    const alreadyLoaded = this.templates.find(
      (template) => template.name === name,
    );
    if (alreadyLoaded) {
      return this;
    }

    const newTemplate = Template.import(location, name);

    if (!this.isPresTemplate(newTemplate)) {
      this.rootTemplate = newTemplate;
    } else {
      this.templates.push(newTemplate);
    }

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
      templateCreationId.push({
        name: template.name,
        slides: creationIds,
      });
    }
    return templateCreationId;
  }

  public modify(cb: ModifyXmlCallback): this {
    this.modifyPresentation.push(cb);
    return this;
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
    slideIdentifier: SourceSlideIdentifier,
    callback?: (slide: Slide) => void,
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

    if (callback !== undefined) {
      newSlide.root = this;
      callback(newSlide);
    }

    this.rootTemplate.slides.push(newSlide);

    return this;
  }

  /**
   * WIP: copy and modify a master from template to output
   * @param name
   * @param masterNumber
   * @param callback
   */
  public addMaster(
    name: string,
    masterNumber: number,
    callback?: (slide: Slide) => void,
  ): this {
    const template = this.getTemplate(name);

    const newMaster = new Master({
      presentation: this,
      template,
      masterNumber,
    });

    // this.rootTemplate.slides.push(newMaster);

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
    await this.writeSlides();
    await this.normalizePresentation();
    await this.applyModifyPresentationCallbacks();

    const rootArchive = await this.rootTemplate.archive;

    const options: JSZip.JSZipGeneratorOptions<'nodebuffer'> = {
      type: 'nodebuffer',
    };

    if (this.params.compression > 0) {
      options.compression = 'DEFLATE';
      options.compressionOptions = {
        level: this.params.compression,
      };
    }

    const content = await rootArchive.generateAsync(options);

    return FileHelper.writeOutputFile(
      this.getLocation(location, 'output'),
      content,
      this,
    );
  }

  /**
   * Write all slides into archive.
   */
  public async writeSlides(): Promise<void> {
    await this.rootTemplate.countExistingSlides();
    this.status.max = this.rootTemplate.slides.length;

    for (const slide of this.rootTemplate.slides) {
      await this.rootTemplate.appendSlide(slide);
    }

    if (this.params.removeExistingSlides) {
      await this.rootTemplate.truncate();
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

    if (this.params.removeExistingSlides) {
      this.modify(ModifyPresentationHelper.removeUnusedFiles);
    }

    this.modify(ModifyPresentationHelper.removedUnusedImages);
    this.modify(ModifyPresentationHelper.removeUnusedContentTypes);
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
          vd('No file matches "' + location + '"');
          vd('@templateDir: ' + this.templateDir);
          vd('@templateFallbackDir: ' + this.templateFallbackDir);
        }
        break;
      case 'output':
        return this.outputDir + location;
      default:
        return location;
    }
  }
}
