import Automizer, { modify, read, XmlHelper } from '../src/index';
import { vd } from '../src/helper/general-helper';
import { XmlSlideHelper } from '../src/helper/xml-slide-helper';
import { ElementInfo } from '../src/types/xml-types';

test('read and re-use shape info, e.g. shape coordinates', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithShapes.pptx`, 'shapes')
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

  pres
    .addSlide('shapes', 2, async (slide) => {
      // Pick a shape and buffer the info object
      info.set(slide.getElement('Star'));

      // Read another shape and print its text fragments:
      const eleInfo = await slide.getElement('Cloud');
      console.log(eleInfo.getText());
    })
    .addSlide('charts', 1, async (slide) => {
      const eleInfo = await info.get();

      // Dump element xml from previous slide:
      // XmlHelper.dump(eleInfo.getXmlElement())

      slide.modifyElement('BarsStacked', [
        modify.setPosition({
          x: 1000000,
          h: eleInfo.position.cx,
          w: eleInfo.position.cy,
        }),
      ]);
    });

  await pres.write(`read-shape-info.test.pptx`);
});
