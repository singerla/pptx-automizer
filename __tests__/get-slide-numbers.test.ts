import Automizer, { modify } from '../src/index';

test('create presentation, add all slides from template using getAllSlideNumbers.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`threeSlidePres.pptx`);

  const allSlideNumbers = await pres.getTemplate('threeSlidePres.pptx').getAllSlideNumbers();

  for (const slideNumber of allSlideNumbers) {
    pres.addSlide('threeSlidePres.pptx', slideNumber);
  }

  const result = await pres.write(`add-all-slides.test.pptx`);

  // Add an assertion to check for slide numbers 1, 2, and 3.
  expect(allSlideNumbers).toEqual([1, 2, 3]);
});
