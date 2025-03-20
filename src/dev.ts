import Automizer from './index';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  const automizer = new Automizer({
    templateDir,
    outputDir,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'emptySlide')
    .load(`SlideWithMedia.pptx`, 'media');

  pres.addSlide('emptySlide', 1, (slide) => {
    slide.addElement('media', 1, 'audio');
  });

  pres.write(`testMedia.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
