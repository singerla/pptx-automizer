import Automizer from '../src/automizer';

test('load empty presentation and append a slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer.loadRoot(`EmptyTemplate.pptx`)
    .load(`SlideWithNotes.pptx`, 'notes');

  pres.addSlide('notes', 1);

  const result = await pres.write(`add-slide-to-empty.test.pptx`);

  expect(result.slides).toBe(1);
});
