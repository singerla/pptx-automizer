import Automizer from '../src/automizer';

test('create presentation and append diagrams', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithDiagrams.pptx`, 'diagrams');

  pres.addSlide('diagrams', 1);
  pres.addSlide('diagrams', 2);

  pres.addSlide('empty', 1, (slide) => {
    slide.addElement('diagrams', 1, 'MatrixDiagram')
  });

  const result = await pres.write(`add-slide-diagrams.test.pptx`);

  expect(result.slides).toBe(4);
});
