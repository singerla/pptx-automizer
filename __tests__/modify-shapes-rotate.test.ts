import Automizer, { modify } from '../src/index';

test('Add and rotate a shape.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes');

  const result = await pres
    .addSlide('shapes', 2, (slide) => {
      slide.modifyElement('Drum', [modify.rotateShape(45)]);
      slide.modifyElement('Cloud', [modify.rotateShape(-45)]);
      slide.modifyElement('Arrow', [modify.rotateShape(180)]);
    })
    .write(`modify-shapes-rotate.test.pptx`);

  expect(result.slides).toBe(2);
});
