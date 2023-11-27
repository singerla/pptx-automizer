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
    slide.generate((pptxGenJSSlide, objectName) => {
      pptxGenJSSlide.addText('Test', {
        x: 1,
        y: 1,
        color: '363636',
        objectName,
      });
    }, 'custom object name');
  });

  const result = await pres.write(`generate-pptxgenjs-text.test.pptx`);

  expect(result.slides).toBe(2);
});
