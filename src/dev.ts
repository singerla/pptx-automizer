import Automizer, { ModifyImageHelper, ModifyShapeHelper } from './index';
import { CmToDxa } from './helper/modify-helper';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    mediaDir: `${__dirname}/../__tests__/media`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .loadMedia(`feather.png`)
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`SlideWithImages.pptx`, 'images');

  await pres
    .addSlide('shapes', 1, (slide) => {
      slide.addElement('images', 2, 'imagePNGduotone', [
        ModifyShapeHelper.setPosition({
          w: CmToDxa(5),
          h: CmToDxa(5),
        }),
        ModifyImageHelper.setRelationTarget('feather.png'),
        ModifyImageHelper.setDuotoneFill({
          tint: 100000,
          color: {
            type: 'srgbClr',
            value: 'ff850c',
          },
        }),
      ]);
    })
    .write(`modify-shapes.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
