import Automizer from '../src/automizer';

test('create presentation and append slides with notes', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithNotes.pptx`, 'notes');

  pres.addSlide('notes', 1);

  const result = await pres.write(`add-slide-notes.test.pptx`);

  expect(result.slides).toBe(2);
});
