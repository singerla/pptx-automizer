import Automizer, { modify, XmlHelper } from './index';
import { vd } from './helper/general-helper';
import * as fs from 'fs';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  // Step 1: Create a pptx with images and a chart inside.
  // The chart is modified by pptx-automizer

  const automizer = new Automizer({
    templateDir,
    outputDir,
    verbosity: 2,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`SlideWithShapes.pptx`);

  pres.addSlide("SlideWithShapes.pptx", 3, async (slide) => {
    const infoTopLevel = await slide.getElement('TopLevelGroup')
    const groupInfoTopLevel = infoTopLevel.getGroupInfo()

  })

  await pres.write(`modify-multi-text.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
