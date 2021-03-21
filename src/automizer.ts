import Template from './classes/template';
import Slide from './classes/slide';
import FileHelper from './helper/file';
import { AutomizerParams, AutomizerSummary } from './types/types';
import { IPresentationProps } from './interfaces/ipresentation-props';
import { PresTemplate } from './interfaces/pres-template';
import { RootPresTemplate } from './interfaces/root-pres-template';

export default class Automizer implements IPresentationProps {
  rootTemplate: RootPresTemplate;
  templates: PresTemplate[];
  templateDir: string;
  outputDir: string;
  timer: number;
  params: AutomizerParams;

  /**
   * Parameters for Automizer constructor.
   * @param params
   */
  constructor(params?: AutomizerParams) {
    this.templates = [];
    this.params = params;

    this.templateDir = (params?.templateDir) ? params.templateDir + '/' : '';
    this.outputDir = (params?.outputDir) ? params.outputDir + '/' : '';

    this.timer = Date.now();
  }

  /**
   * Load a pptx file and set it as root template.
   * @param {string} location - Filename or path to the template. Will be prefixed with 'templateDir'
   * @return {Automizer} Instance of Automizer
   */
  public loadRoot(location: string): this {
    return this.loadTemplate(location);
  }

  /**
   * Load a template pptx file.
   * @param {string} location - Filename or path to the template. Will be prefixed with 'templateDir'
   * @param {string} name - Optional: A short name for the template. If skipped, the template will be named by its location.
   * @return {Automizer} Instance of Automizer
   */
  public load(location: string, name?: string): this {
    name = (name === undefined) ? location : name;
    return this.loadTemplate(location, name);
  }

  public loadTemplate(location: string, name?: string): this {
    location = this.getLocation(location, 'template');

    const newTemplate = Template.import(location, name);

    if (!this.isPresTemplate(newTemplate)) {
      this.rootTemplate = newTemplate;
    } else {
      this.templates.push(newTemplate);
    }

    return this;
  }

  isPresTemplate(template: PresTemplate | RootPresTemplate): template is PresTemplate {
    return 'name' in template;
  }

  /**
   * Find imported template by given name and return a certain slide by number.
   * @param {string} name - Name of template; must be imported by Automizer.importTemplate()
   * @param {number} slideNumber - Number of slide in template presentation
   * @param callback
   * @return {Automizer} Instance of Automizer
   */
  public addSlide(name: string, slideNumber: number, callback?: (slide: Slide) => void): this {
    if (this.rootTemplate === undefined) {
      throw new Error('You have to set a root template first.');
    }

    const template = this.template(name);

    const newSlide = new Slide({
      presentation: this,
      template,
      number: slideNumber
    });

    if (callback !== undefined) {
      newSlide.root = this;
      callback(newSlide);
    }

    this.rootTemplate.slides.push(newSlide);

    return this;
  }

  public template(name: string): PresTemplate {
    const template = this.templates.find(t => t.name === name);
    if (template === undefined) {
      throw new Error(`Template not found: ${name}`);
    }
    return template;
  }

  public getLocation(location: string, type?: string): string {
    switch (type) {
      case 'template':
        return this.templateDir + location;
      case 'output':
        return this.outputDir + location;
      default:
        return location;
    }
  }

  async write(location: string): Promise<AutomizerSummary> {
    const rootArchive = await this.rootTemplate.archive;

    for (const slide of this.rootTemplate.slides) {
      await this.rootTemplate.appendSlide(slide);
    }

    const content = await rootArchive.generateAsync({type: 'nodebuffer'});

    return FileHelper.writeOutputFile(this.getLocation(location, 'output'), content, this);
  }
}
