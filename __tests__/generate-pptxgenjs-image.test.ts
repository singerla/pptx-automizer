import Automizer from '../src/automizer';
import { ChartData, modify } from '../src';

test('insert an image with pptxgenjs on a template slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add image from file:
    slide.generate((pptxGenJSSlide, objectName) => {
      pptxGenJSSlide.addImage({
        path: `${__dirname}/images/test.png`,
        x: 1,
        y: 2,
        objectName,
      });
    });
  });

  const result = await pres.write(`generate-pptxgenjs-image.test.pptx`);

  expect(result.images).toBe(1);
});
