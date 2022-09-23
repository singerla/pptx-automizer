import Automizer, { ChartData, modify, TableRow, TableRowStyle } from './index';
import { vd } from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
});

const run = async () => {
  const pres = automizer
    .loadRoot(`bug/RootTemplate.pptx`)
    .load(`bug/Slides.pptx`, 'slides')
    .load(`bug/Library - Icons.pptx`, 'image');

  const result = await pres
    .addSlide('slides', 6, (slide) => {
      slide.addElement('image', 1, 'Tschechien');
    })
    .write(`add-image-bug.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
