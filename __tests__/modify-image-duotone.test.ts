import Automizer from '../src/automizer';
import { ModifyImageHelper, ModifyShapeHelper } from '../src';
import { CmToDxa } from '../src/helper/modify-helper';

test('Add image and set duotone fill', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    removeExistingSlides: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`SlideWithImages.pptx`, 'images');

  pres.addSlide('shapes', 1, (slide) => {
    slide.addElement('images', 2, 'imagePNGduotone', [
      ModifyImageHelper.setDuotoneFill({
        tint: 100000,
        color: {
          type: 'srgbClr',
          value: 'ff0000',
        },
      }),
    ]);
  });

  const result = await pres.write(`modify-image-duotone.test.pptx`);

  // Expect cord loop to turn red by duotone overlay

  expect(result.images).toBe(1);
});
