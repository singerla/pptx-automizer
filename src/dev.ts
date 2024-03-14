import Automizer, { XmlHelper } from './index';
import ModifyBackgroundHelper from './helper/modify-background-helper';
import { vd } from './helper/general-helper';
import ModifyImageHelper from './helper/modify-image-helper';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    autoImportSlideMasters: true,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideMasterBackgrounds.pptx`)

    .load('SlideMasterBackgrounds.pptx')
    .loadMedia(`test.png`, `${__dirname}/../__tests__/media`, 'pre_')
    .addMaster(`SlideMasterBackgrounds.pptx`, 2, async (master) => {
      ModifyBackgroundHelper.setRelationTarget(master, 'pre_test.png');
    })
    .addSlide(`SlideMasterBackgrounds.pptx`, 1)
    .write(`SlideMasterBackgroundsOutput.pptx`)
    .then((summary) => {
      console.log(summary);
    });
};

run().catch((error) => {
  console.error(error);
});
