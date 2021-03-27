import Automizer from './index';
import { setChartData } from './helper/modify-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`ChartBarsStackedFmt.pptx`, 'charts');

const data = {
  series: [
    { label: 'series s1' },
    { label: 'series s2' },
    { label: 'series s3' },
  ],
  categories: [
    { label: 'item test r1', values: [10, 16, 12] },
    { label: 'item test r2', values: [12, 18, 15] },
    // { label: 'item test r3', values: [ 14, 12, 11 ] },
    // { label: 'item test r4', values: [ 8, 11, 9 ] },
    // { label: 'item test r5', values: [ 6, 15, 7 ] },
    // { label: 'item test r6', values: [ 16, 16, 9 ] },
  ],
};

pres
  .addSlide('charts', 1, (slide) => {
    slide.modifyElement('BarsStacked', [setChartData(data)]);
  })
  .write(`modify-chart-stacked-bars.test.pptx`)

  .then((result) => {
    console.info(result);
  })
  .catch((error) => {
    console.error(error);
  });
