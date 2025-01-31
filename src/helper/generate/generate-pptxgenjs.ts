import { randomUUID } from 'crypto';
import PptxGenJS from 'pptxgenjs';
import fs from 'fs';
import { ISlide } from '../../interfaces/islide';
import Automizer from '../../automizer';
import { GenerateElements, SupportedPptxGenJSSlide } from '../../types/types';
import { IGenerator } from '../../interfaces/igenerator';

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
      generateElement.tmpSlideNumber = this.countSlides;

      const addedObjects = <string[]>[];

      generateElement.callback(
        this.supportedSlideItems(pgenSlide, generateElement, addedObjects),
        this.generator,
      );

      addedObjects.forEach((addedObjectName) => {
        slide.addElement(this.tmpFile, this.countSlides, addedObjectName);
      });
    });
  }

  supportedSlideItems = (
    pgenSlide: PptxGenJS.Slide,
    generateElement: GenerateElements,
    addedObjects: string[],
  ): SupportedPptxGenJSSlide => {
    return {
      addChart: (type, data, options) => {
        const objectName = this.generateObjectName(
          generateElement,
          addedObjects,
        );
        pgenSlide.addChart(type, data, this.getOptions(options, objectName));
      },
      addImage: (options) => {
        const objectName = this.generateObjectName(
          generateElement,
          addedObjects,
        );
        pgenSlide.addImage(this.getOptions(options, objectName));
      },
      addShape: (shapeName, options?) => {
        const objectName = this.generateObjectName(
          generateElement,
          addedObjects,
        );
        pgenSlide.addShape(shapeName, this.getOptions(options, objectName));
      },
      addTable: (tableRows, options?) => {
        const objectName = this.generateObjectName(
          generateElement,
          addedObjects,
        );
        pgenSlide.addTable(tableRows, this.getOptions(options, objectName));
      },
      addText: (text, options?) => {
        const objectName = this.generateObjectName(
          generateElement,
          addedObjects,
        );
        pgenSlide.addText(text, this.getOptions(options, objectName));
      },
    };
  };

  generateObjectName(generateElement, addedObjects: string[]): string {
    const objectName =
      (generateElement.objectName ? generateElement.objectName + '-' : '') +
      randomUUID();
    addedObjects.push(objectName);
    return objectName;
  }

  getOptions = (options, objectName: string) => {
    options = options || {};
    return {
      ...options,
      objectName,
    };
  };

  appendPptxGenSlide(): PptxGenJS.Slide {
    return this.generator.addSlide();
  }

  async cleanup() {
    if (this.countSlides > 0) {
      fs.unlinkSync(this.automizer.templateDir + '/' + this.tmpFile);
    }
  }
}
