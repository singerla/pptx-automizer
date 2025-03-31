import Automizer from '../src/automizer';

test('create presentation and append slides with images', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SVGImages.pptx`, 'images');

  pres.addSlide('empty', 1, (slide) => {
    slide.addElement('images', 1, 'Heart');
    slide.addElement('images', 1, 'Leaf');
  });

  // Test processed related svg content on added slides
  pres.addSlide('images', 1);

  const result = await pres.write(`add-svg-images.test.pptx`);

  expect(result.images).toBe(10);
});
