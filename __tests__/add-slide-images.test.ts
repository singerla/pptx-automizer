import Automizer from '../src/automizer';

test('create presentation and append slides with images', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
    .load(`SlideWithImages.pptx`, 'images');

  pres.addSlide('images', 1);
  pres.addSlide('images', 2);

  const result = await pres.write(`add-slide-images.test.pptx`);

  expect(result.images).toBe(5);
});
