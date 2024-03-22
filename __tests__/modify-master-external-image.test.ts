import Automizer from '../src/automizer';
import { ModifyImageHelper, ModifyShapeHelper } from '../src';
import { CmToDxa } from '../src/helper/modify-helper';

test('Load external media, modify image target on slide master', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    mediaDir: `${__dirname}/../__tests__/media`,
    removeExistingSlides: true,
    cleanup: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplateWithImages.pptx`)
    .loadMedia([`test.png`])
    .load(`RootTemplateWithImages.pptx`, 'base');

  pres.addMaster('base', 1, (master) => {
    master.modifyElement('masterImagePNG', [
      ModifyImageHelper.setRelationTarget('test.png'),
    ]);
  });

  // Expect imported slide master (#2) to have swapped (left top) background image
  const result = await pres.write(`modify-master-add-external-image.test.pptx`);

  expect(result.images).toBe(8);
});
