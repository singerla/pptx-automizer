import Automizer from '../src/index';

test('Add and modify hyperlinks', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithLink.pptx`, 'link');

  const result = await pres
    .addSlide('link', 1, (slide) => {})
    .addSlide('empty', 1, (slide) => {
      slide.addElement('link', 1, 'Link');
    })
    .write(`modify-hyperlink.test.pptx`);

  expect(result.slides).toBe(3);
});
