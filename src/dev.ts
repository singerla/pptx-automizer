import Automizer, { modify, XmlHelper } from './index';
import { vd } from './helper/general-helper';
import * as fs from 'fs';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  const automizer = new Automizer({
    templateDir,
    outputDir,
    verbosity: 2,
    removeExistingSlides: true,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`SlideWithDiagram.pptx`);

  pres.addSlide("SlideWithDiagram.pptx", 1, async (slide) => {

  })

  await pres.write(`add-diagram.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
