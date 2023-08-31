import Automizer, { modify } from './index';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates/customer`,
    outputDir: `${__dirname}/../__tests__/pptx-templates/customer`,
    autoImportSlideMasters: true,
    showIntegrityInfo: true,
    assertRelatedContents: true,
    useCreationIds: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`Chapter 8 - Online Shopping.pptx`, 'shapes');

  const result = await pres
    .addSlide('shapes', 1, (slide) => {})
    .write(`test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
