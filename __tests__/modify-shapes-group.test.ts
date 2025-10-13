import Automizer, { modify } from '../src/index';

test('modify of grouped shapes', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes');

  const result = await pres
    .addSlide('shapes', 3, (slide) => {
      // slide.modifyElement('Arrow', modify.setText('stays in group'));
      slide.addElement('shapes', 3, 'Arrow', modify.setText('stays in group'));
    })
    .write(`modify-shapes-group.test.pptx`);

  expect(result.slides).toBe(2);
});
