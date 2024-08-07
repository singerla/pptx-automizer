import Automizer from '../src/index';
import { ElementInfo } from '../src/types/xml-types';

test('read alt text info', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithImages.pptx`, 'images')
    .load(`ChartBarsStacked.pptx`, 'charts');

  // A simple helper to get/set ElementInfo
  const info = {
    element: <Promise<ElementInfo>>{},
    set: (elementInfo: Promise<ElementInfo>) => {
      info.element = elementInfo;
    },
    get: async (): Promise<ElementInfo> => {
      return info.element;
    },
  };
let altText = '';
  pres
    .addSlide('images', 1, async (slide) => {
      // Read another shape and print its text fragments:
      const eleInfo = await slide.getElement('Grafik 5');
      altText = eleInfo.altText;
    });

  await pres.write(`read-alt-text.test.pptx`);

  expect(altText).toBe('picture');
});
