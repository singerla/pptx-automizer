import Automizer from '../src/automizer';
import { ModifyTextHelper } from '../src';
import ModifyBackgroundHelper from '../src/helper/modify-background-helper';

test('Auto-import source slideLayout and -master and modify background color', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    autoImportSlideMasters: true,
  });

  const pres = await automizer
    .loadRoot(`EmptyTemplate.pptx`)
    .load('SlideMasterBackgrounds.pptx')
    .addMaster(`SlideMasterBackgrounds.pptx`, 1, async (master) => {
      master.modify(
        ModifyBackgroundHelper.setSolidFill({
          type: 'srgbClr',
          value: 'aaccbb',
        }),
      );
    })
    .addSlide(`SlideMasterBackgrounds.pptx`, 3)
    .write(`modify-master-background-color.test.pptx`);

  expect(pres.masters).toBe(2);
});
