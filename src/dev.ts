import Automizer from './index';
import { vd } from './helper/general-helper';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    removeExistingSlides: true,
  });

  let pres = automizer
    .loadRoot(`SlidesWithoutCreationIds.pptx`)
    .load(`SlideWithCharts.pptx`, 'noCreationId');

  pres.addSlide('noCreationId', 1, async (slide) => {
    // const elements = await slide.getAllElements();
    // const textElements = await slide.getAllTextElementIds();
    vd(await slide.getDimensions());
  });

  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
