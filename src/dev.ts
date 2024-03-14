import Automizer from './index';
import ModifyBackgroundHelper from './helper/modify-background-helper';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    autoImportSlideMasters: true,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideMasterBackgrounds.pptx`);

  pres.addMaster(`SlideMasterBackgrounds.pptx`, 1, async (master) => {
    master.modify(
      ModifyBackgroundHelper.setSolidFill({
        type: 'srgbClr',
        value: 'aaccbb',
      }),
    );
  });

  pres.addSlide(`SlideMasterBackgrounds.pptx`, 3, async (slide) => {
    console.log('test');
  });

  pres.write(`SlideMasterBackgroundsOutput.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
