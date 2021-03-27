import Automizer from '../src/automizer';

test('create presentation and add some single images', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithImages.pptx`, 'images');

  const result = await pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('images', 2, 'imageJPG');
      slide.addElement('images', 2, 'imagePNG');
    })
    .write(`add-single-images.test.pptx`);

  expect(result.slides).toBe(2);
});
