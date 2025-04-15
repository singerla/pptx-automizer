import Automizer from '../src/automizer';

test('create presentation and add basic slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    archiveType: {
      mode: 'jszip',
      decodeText: true,
    },
  });
  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes');

  for (let i = 0; i <= 10; i++) {
    pres.addSlide('shapes', 1);
  }

  await pres.write(`create-presentation.test.pptx`);

  expect(pres).toBeInstanceOf(Automizer);
});
