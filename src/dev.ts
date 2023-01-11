import Automizer, { ChartData, modify } from './index';
import { vd } from './helper/general-helper';
import { contentTracker } from './helper/content-tracker';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
  compression: 9,
});

const run = async () => {
  const pres = automizer
    .loadRoot(`ChartBarsStacked.pptx`)
    .load(`RootTemplate.pptx`, 'root')
    .load(`ChartBarsStacked.pptx`, 'charts');

  const data = {
    series: [
      { label: 'series s1' },
      { label: 'series s2' },
      { label: 'series s3' },
      { label: 'series s4' },
    ],
    categories: [
      { label: 'item test r1', values: [10, 16, 12, 15] },
      { label: 'item test r2', values: [12, 18, 15, 15] },
      { label: 'item test r3', values: [14, 12, 11, 15] },
      { label: 'item test r4', values: [8, 11, 9, 15] },
      { label: 'item test r5', values: [6, 15, 7, 15] },
      { label: 'item test r6', values: [16, 16, 9, 3] },
    ],
  };

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(data)]);
      slide.addElement('charts', 1, 'BarsStacked', [modify.setChartData(data)]);
    })
    .addSlide('root', 1, (slide) => {
      slide.addElement('charts', 1, 'BarsStacked', [modify.setChartData(data)]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .write(`create-presentation-content-tracker.test.pptx`);

  // vd(pres.rootTemplate.content);
};

run().catch((error) => {
  console.error(error);
});
