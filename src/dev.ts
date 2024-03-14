import Automizer from './index';
import { vd } from './helper/general-helper';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    autoImportSlideMasters: true,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideMasterBackgrounds.pptx`);

  pres.addSlide(`SlideMasterBackgrounds.pptx`, 2, async (slide) => {
    console.log('test');
  });

  pres.write(`SlideMasterBackgroundsOutput.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
