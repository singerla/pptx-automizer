import JSZip, { InputType } from 'jszip';

import { FileHelper } from '../helper/file-helper';
import { CountHelper } from '../helper/count';
import { ICounter } from '../interfaces/icounter';
import { ISlide } from '../interfaces/islide';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ITemplate } from '../interfaces/itemplate';

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
   * @type Promise<Buffer>
   */
  file: InputType;

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

  constructor(location: string) {
    this.location = location;
    const file = FileHelper.readFile(location);
    this.archive = FileHelper.extractFileContent((file as unknown) as Buffer);
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
