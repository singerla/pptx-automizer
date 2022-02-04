import Automizer, {ChartData, modify} from './index';
import {vd} from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`TemplateWithMaster.pptx`, 'master');


const run = async () => {
  await pres
    .addMaster('master', 2, (slide) => {

    })
    .addSlide('master', 2, (slide) => {

    })
    .write(`add-master.test.pptx`)

  return pres;
};

run().catch((error) => {
  console.error(error);
});
