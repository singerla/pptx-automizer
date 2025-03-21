import Automizer from '../src/automizer';

test('create presentation and add media shapes', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'emptySlide')
    .load(`SlideWithMedia.pptx`, 'media');

  pres.addSlide('emptySlide', 1, (slide) => {
    slide.addElement('media', 1, 'audio');
  });

  pres.addSlide('emptySlide', 1, (slide) => {
    slide.addElement('media', 2, 'video');
    slide.addElement('media', 1, 'audio');
  });

  // TODO: Process related media content on added slides
  // pres.addSlide('media', 1, (slide) => {
  //   slide.addElement('media', 2, 'video');
  // });

  const result = await pres.write(`add-media.test.pptx`);

  // expect(result.images).toBe(6);
});
