import Automizer from '../src/automizer';
import { ModifyImageHelper, ModifyTextHelper, XmlHelper } from '../src';
import ModifyBackgroundHelper from '../src/helper/modify-background-helper';

test('Auto-import source slideLayout and -master and modify background image', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    autoImportSlideMasters: true,
  });

  const pres = await automizer
    .loadRoot(`EmptyTemplate.pptx`)
    .load('SlideMasterBackgrounds.pptx')
    .loadMedia(`test.png`, `${__dirname}/../__tests__/media`, 'pre_')
    .addMaster(`SlideMasterBackgrounds.pptx`, 2, async (master) => {
      ModifyBackgroundHelper.setRelationTarget(master, 'pre_test.png');
    })
    .addSlide(`SlideMasterBackgrounds.pptx`, 1)
    .addSlide(`SlideMasterBackgrounds.pptx`, 1)
    .write(`modify-master-background-image.test.pptx`);

  expect(pres.masters).toBe(2);
});
