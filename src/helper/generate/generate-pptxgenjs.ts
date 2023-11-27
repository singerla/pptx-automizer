import { randomUUID } from 'crypto';
import PptxGenJS from 'pptxgenjs';
import fs from 'fs';
import { ISlide } from '../../interfaces/islide';
import Automizer from '../../automizer';
import { GenerateElements } from '../../types/types';
import { IGenerator } from '../../interfaces/igenerator';
import { vd } from '../general-helper';

export default class GeneratePptxGenJs implements IGenerator {
  tmpFile: string;
  slides: ISlide[];
  generator: PptxGenJS;
  automizer: Automizer;
  countSlides: number = 0;

  constructor(automizer: Automizer, slides: ISlide[]) {
    this.automizer = automizer;
    this.slides = slides;
  }

  create(): this {
    this.generator = new PptxGenJS();
    return this;
  }

  async generateSlides(): Promise<void> {
    this.tmpFile = randomUUID() + '.pptx';
    for (const slide of this.slides) {
      const generate = slide.getGeneratedElements();
      if (generate.length) {
        this.countSlides++;
        this.addElements(generate, this.appendPptxGenSlide(), slide);
      }
    }

    if (this.countSlides > 0) {
      await this.generator.writeFile({
        fileName: this.automizer.templateDir + '/' + this.tmpFile,
      });
      this.automizer.load(this.tmpFile);
    }
  }

  addElements(
    generate: GenerateElements[],
    pgenSlide: PptxGenJS.Slide,
    slide: ISlide,
  ) {
    generate.forEach((generateElement) => {
      generateElement.objectName = generateElement.objectName || randomUUID();
      generateElement.tmpSlideNumber = this.countSlides;
      generateElement.callback(
        pgenSlide,
        generateElement.objectName,
        this.generator,
      );
      slide.addElement(
        this.tmpFile,
        this.countSlides,
        generateElement.objectName,
      );
    });
  }

  appendPptxGenSlide(): PptxGenJS.Slide {
    return this.generator.addSlide();
  }

  async cleanup() {
    if (this.countSlides > 0) {
      fs.unlinkSync(this.automizer.templateDir + '/' + this.tmpFile);
    }
  }
}
