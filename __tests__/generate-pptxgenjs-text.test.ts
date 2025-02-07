import Automizer from '../src/automizer';
import { ChartData, modify } from '../src';

test('insert a textbox with pptxgenjs on a template slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add text from scratch:
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addText('Test 1', {
        x: 1,
        y: 1,
        h: 5,
        w: 10,
        color: '363636',
      });
    }, 'custom object name');
  });

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add text from scratch:
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addText('Test 2', {
        x: 1,
        y: 1,
        h: 5,
        w: 10,
        color: '363636',
      });
      pptxGenJSSlide.addText('Test 3', {
        x: 1,
        y: 1,
        h: 5,
        w: 10,
        color: '363636',
      });
    }, 'custom object name');
  });

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add text from scratch:
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addText('Test 4', {
        x: 1,
        y: 1,
        h: 5,
        w: 10,
        color: '363636',
      });
      pptxGenJSSlide.addText('Test 5', {
        x: 1,
        y: 1,
        h: 5,
        w: 10,
        color: '363636',
      });
      pptxGenJSSlide.addText('Test 6', {
        x: 1,
        y: 1,
        h: 5,
        w: 10,
        color: '363636',
      });
    }, 'custom object name');
  });

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add text from scratch:
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addText('Test 7', {
        x: 1,
        y: 6,
        h: 5,
        w: 10,
        color: '363636',
      });
      pptxGenJSSlide.addText('Test 8', {
        x: 1,
        y: 3,
        h: 5,
        w: 10,
        color: '363636',
      });
      pptxGenJSSlide.addText('Test 9', {
        x: 1,
        y: 2,
        h: 5,
        w: 10,
        color: '363636',
      });
    }, 'custom object name 1243453');

    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addImage({
        path: `${__dirname}/media/test.png`,
        x: 1,
        y: 2,
      });
      pptxGenJSSlide.addImage({
        path: `${__dirname}/images/test.svg`,
        x: 4,
        y: 2,
      });
    }, 'custom object name 123');
  });

  const result = await pres.write(`generate-pptxgenjs-text.test.pptx`);

  // expect(result.slides).toBe(2);
});
