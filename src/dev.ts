import Automizer from './index';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  const automizer = new Automizer({
    templateDir,
    outputDir,
    autoImportSlideMasters: true,
    cleanupPlaceholders: true,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlidePlaceholders.pptx`, 'placeholder')
    .load(`EmptySlide.pptx`, 'emptySlide');

  pres.addSlide('emptySlide', 1, (slide) => {
    slide.addElement('placeholder', 1, 'Titel 4');
  });

  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
