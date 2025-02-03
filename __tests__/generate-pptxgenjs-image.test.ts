import Automizer from '../src/automizer';

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
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addImage({
        path: `${__dirname}/media/test.png`,
        x: 1,
        y: 2,
      });
      pptxGenJSSlide.addImage({
        path: `${__dirname}/images/test.svg`,
        x: 1,
        y: 2,
      });
    });
  });

  const result = await pres.write(`generate-pptxgenjs-image.test.pptx`);

  expect(result.images).toBe(3);
});
