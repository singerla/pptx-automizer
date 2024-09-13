import Automizer from '../src/index';

test('read alt text from image', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images');

  let altText = '';
  pres.addSlide('images', 1, async (slide) => {
    // Read alt text from an image:
    const eleInfo = await slide.getElement('Grafik 5');
    altText = eleInfo.getAltText();
  });

  await pres.write(`read-alt-text.test.pptx`);

  expect(altText).toBe('picture');
});
