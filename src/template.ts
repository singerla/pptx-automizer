import {
  ISlide, ITemplate, PresTemplate, RootPresTemplate, ICounter
} from './definitions/app';

import FileHelper from './helper/file';
import JSZip from 'jszip';
import CountHelper from './helper/count';

class Template implements ITemplate {
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
   * @type Promise<Buffer>
   */
  file: Promise<Buffer>;

  /**
   * this.file will be passed to JSZip
   * @type Promise<JSZip>
   */
  archive: Promise<JSZip>;

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

  constructor(location: string, name?: string) {
    this.location = location;
    this.file = FileHelper.readFile(location);
    this.archive = FileHelper.extractFileContent(this.file);
  }

  static import(location: string, name?: string): PresTemplate | RootPresTemplate {
    let newTemplate: PresTemplate | RootPresTemplate;

    if (name) {
      newTemplate = <PresTemplate>new Template(location, name);
      newTemplate.name = name;
    } else {
      newTemplate = <RootPresTemplate><unknown>new Template(location);
      newTemplate.slides = [];
      newTemplate.counter = [
        new CountHelper('slides', newTemplate),
        new CountHelper('charts', newTemplate),
        new CountHelper('images', newTemplate)
      ];
    }

    return newTemplate;
  }

  async appendSlide(slide: ISlide): Promise<void> {
    if (this.counter[0].get() === undefined) {
      await this.initializeCounter();
    }

    await slide.append(this);
  }

  async initializeCounter(): Promise<void> {
    for (let i in this.counter) {
      await this.counter[i].set();
    }
  }

  incrementCounter(name: string): number {
    return CountHelper.increment(name, this.counter);
  }

  count(name: string): number {
    return CountHelper.count(name, this.counter);
  }
}


export default Template;
