import { randomUUID } from 'crypto';
import PptxGenJS from 'pptxgenjs';
import { ISlide } from '../../interfaces/islide';
import Automizer from '../../automizer';
import { AutomizerFile, GenerateElements } from '../../types/types';
import { IGenerator } from '../../interfaces/igenerator';
import { IPptxGenJSSlide } from '../../interfaces/ipptxgenjs-slide';
import fs from 'fs';

/**
 * Using pptxGenJs on an automizer ISlide will create a temporary pptx template
 * and auto-import the generated shapes to the right place on the output slides.
 */
export default class GeneratePptxGenJs implements IGenerator {
  tmpFile: string;
  slides: ISlide[];
  generator: PptxGenJS;
  automizer: Automizer;
  countSlides: number = 0;

  constructor(automizer: Automizer, slides: ISlide[]) {
    this.automizer = automizer;
    this.slides = slides;
    this.create();
  }

  create() {
    if (this.automizer.params.pptxGenJs) {
      // Use a customized pptxGenJs instance
      this.generator = this.automizer.params.pptxGenJs;
    } else {
      // Or the installed version
      this.generator = new PptxGenJS();
    }
  }

  async generateSlides(): Promise<void> {
    this.tmpFile = randomUUID() + '.pptx';

    for (const slide of this.slides) {
      const generate = slide.getGeneratedElements();

      if (generate.length) {
        this.countSlides++;
        const pgenSlide = this.appendPptxGenSlide();
        await this.generateElements(generate, pgenSlide, this.countSlides);
      }
    }

    for (const slide of this.slides) {
      const generate = slide.getGeneratedElements();
      if (generate.length) {
        this.addElements(generate, slide);
      }
    }

    if (this.countSlides > 0) {
      const data = (await this.generator.stream()) as AutomizerFile;
      this.automizer.load(data, this.tmpFile);

      // await this.generator.writeFile({
      //   fileName: this.automizer.templateDir + '/' + this.tmpFile,
      // });
    }
  }

  async generateElements(
    generate: GenerateElements[],
    pgenSlide,
    tmpSlideNumber,
  ): Promise<void> {
    for (const generateElement of generate) {
      generateElement.tmpSlideNumber = tmpSlideNumber;
      const addedObjects = <string[]>[];
      await generateElement.callback(
        this.addSlideItems(pgenSlide, generateElement, addedObjects),
        this.generator,
      );
      generateElement.addedObjects = [...addedObjects];
    }
  }

  addElements(generate: GenerateElements[], slide: ISlide) {
    generate.forEach((generateElement) => {
      generateElement.addedObjects.forEach((addedObjectName) => {
        slide.addElement(
          this.tmpFile,
          generateElement.tmpSlideNumber,
          addedObjectName,
        );
      });
    });
  }

  /**
   * This is a wrapper around supported pptxGenJS slide item types.
   * It is required to create a unique objectName and find the generated
   * shapes by object name later.
   *
   * @param pgenSlide
   * @param generateElement
   * @param addedObjects
   */
  addSlideItems = (
    pgenSlide: PptxGenJS.Slide,
    generateElement: GenerateElements,
    addedObjects: string[],
  ): IPptxGenJSSlide => {
    const getObjectName = () => {
      return this.generateObjectName(generateElement, addedObjects);
    };
    return {
      addChart: (type, data, options) => {
        pgenSlide.addChart(
          type,
          data,
          this.getOptions(options, getObjectName()),
        );
      },
      addImage: (options) => {
        pgenSlide.addImage(this.getOptions(options, getObjectName()));
      },
      addShape: (shapeName, options?) => {
        pgenSlide.addShape(
          shapeName,
          this.getOptions(options, getObjectName()),
        );
      },
      addTable: (tableRows, options?) => {
        pgenSlide.addTable(
          tableRows,
          this.getOptions(options, getObjectName()),
        );
      },
      addText: (text, options?) => {
        pgenSlide.addText(text, this.getOptions(options, getObjectName()));
      },
    };
  };

  generateObjectName(
    generateElement: GenerateElements,
    addedObjects: string[],
  ): string {
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
    // if (this.countSlides > 0) {
    //   fs.unlinkSync(this.automizer.templateDir + '/' + this.tmpFile);
    // }
  }
}
