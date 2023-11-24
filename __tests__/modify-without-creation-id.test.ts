import Automizer, { modify } from '../src/index';
import { vd } from '../src/helper/general-helper';

test('load template without slide & element creation ids.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlidesWithoutCreationIds.pptx`, 'noCreationId');

  const result = await pres
    .addSlide('noCreationId', 1, async (slide) => {
      const elements = await slide.getAllElements();
      expect(elements.length).toBe(3);

      const textElements = await slide.getAllTextElementIds();
      expect(textElements.length).toBe(2);
    })
    .write(`modify-without-creation-id.test.pptx`);
});
