import Automizer from '../src/automizer';

test('create presentation, add slides, remove elements and add one.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`SlideWithImages.pptx`, 'images');

  const result = await pres
    .addSlide('charts', 2, (slide) => {
      slide.removeElement('ColumnChart');
    })
    .addSlide('images', 2, (slide) => {
      slide.removeElement('imageJPG');
      slide.removeElement('Textfeld 5');
      slide.addElement('images', 2, 'imageJPG');
    })
    .write(`remove-existing-elements.test.pptx`);

  expect(result.slides).toBe(3);
});
