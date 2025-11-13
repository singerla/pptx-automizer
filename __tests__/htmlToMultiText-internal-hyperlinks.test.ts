import Automizer, { modify } from '../src/index';

test('create presentation with internal slide hyperlinks using htmlToMultiText', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  // Please note: You need to recalculate the target slide number and
  // refer to the final position, including the existing slides.
  // Internally, we target to `slide${target}.xml`, which might be a different slide.
  // Please make sure that the internally linked slide does exist, otherwise the link will not be active.

  const html = '<body><p>See <a href="4">slide 3</a> for details</p></body>';

  const result = await pres
    .addSlide('TextReplace.pptx', 1)
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .addSlide('TextReplace.pptx', 1)
    .addSlide('TextReplace.pptx', 1)
    .addSlide('TextReplace.pptx', 1)
    .write(`htmlToMultiText-internal-hyperlinks.test.pptx`);

  // Test passes if file is written successfully
});
