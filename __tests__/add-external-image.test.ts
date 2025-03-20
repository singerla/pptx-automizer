import Automizer from '../src/automizer';
import { ModifyImageHelper, ModifyShapeHelper } from '../src';
import { CmToDxa } from '../src/helper/modify-helper';

test('Load external media, add/modify image and set image target', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    mediaDir: `${__dirname}/../__tests__/media`,
    removeExistingSlides: true,
    cleanup: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .loadMedia([`feather.png`, `test.png`, `Dàngerous Dinösaur.png`])
    .loadMedia(`test.png`, `${__dirname}/../__tests__/media`, 'pre_')
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`SlideWithImages.pptx`, 'images');

  pres.addSlide('shapes', 1, (slide) => {
    slide.addElement('images', 2, 'imagePNGduotone', [
      ModifyShapeHelper.setPosition({
        w: CmToDxa(5),
        h: CmToDxa(3),
      }),
      ModifyImageHelper.setRelationTarget('feather.png'),
    ]);
  });

  pres.addSlide('images', 1, (slide) => {
    slide.modifyElement('Grafik 5', [
      ModifyImageHelper.setRelationTarget('pre_test.png'),
    ]);
  });

  pres.addSlide('images', 1, (slide) => {
    slide.modifyElement('Grafik 5', [
      ModifyImageHelper.setRelationTarget('Dàngerous Dinösaur.png'),
    ]);
  });

  const result = await pres.write(`add-external-image.test.pptx`);

  // expect a 5x3cm light-blue duotone feather instead of imagePNG cord loop on page 1
  // expect imagePNG cord loop on page 2 instead of cut tree jpg

  expect(result.images).toBe(5);
});
