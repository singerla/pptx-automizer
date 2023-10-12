import Automizer from '../src/automizer';
import { ModifyImageHelper, ModifyShapeHelper } from '../src';
import { CmToDxa } from '../src/helper/modify-helper';

test('Load external media, add image and set image target', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    mediaDir: `${__dirname}/../__tests__/media`,
    removeExistingSlides: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .loadMedia([`feather.png`, `test.png`])
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`SlideWithImages.pptx`, 'images');

  pres.addSlide('shapes', 1, (slide) => {
    slide.addElement('images', 2, 'imagePNG', [
      ModifyShapeHelper.setPosition({
        w: CmToDxa(5),
        h: CmToDxa(5),
      }),
      ModifyImageHelper.setRelationTarget('test.png'),
    ]);
  });

  const result = await pres.write(`add-external-image.test.pptx`);

  // expect(result.images).toBe(5);
});
