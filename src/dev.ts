import Automizer, { ChartData, modify } from './index';
import { vd } from './helper/general-helper';
import { contentTracker } from './helper/content-tracker';
import ModifyPresentationHelper from './helper/modify-presentation-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
  cleanup: true,
  compression: 5,
});

const run = async () => {
  const pres = automizer
    .loadRoot(`RootTemplateWithImages.pptx`)
    .load(`RootTemplate.pptx`, 'root')
    .load(`SlideWithImages.pptx`, 'images')
    .load(`ChartBarsStacked.pptx`, 'charts');

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
      slide.addElement('charts', 1, 'BarsStacked', [
        modify.setChartData(dataSmaller),
      ]);
    })
    .addSlide('images', 1)
    .addSlide('root', 1, (slide) => {
      slide.addElement('charts', 1, 'BarsStacked', [
        modify.setChartData(dataSmaller),
      ]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.addElement('images', 2, 'imageJPG');
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .write(`create-presentation-content-tracker.test.pptx`);

  // vd(pres.rootTemplate.content);
};

run().catch((error) => {
  console.error(error);
});
