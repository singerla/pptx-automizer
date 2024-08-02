import Automizer from '../src/automizer';
import { vd } from '../src/helper/general-helper';

test('create presentation and append slides with notes', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithNotes.pptx`, 'notes');

  pres.addSlide('notes', 1);

  const creationIds = await pres.setCreationIds();
  // This will print out the first line of a slide note field
  // OR the slide title, if there is no slide note present:
  console.log(creationIds[0].slides[0].info);

  const result = await pres.write(`add-slide-notes.test.pptx`);

  expect(result.slides).toBe(2);
});
