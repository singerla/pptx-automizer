import Automizer, { modify } from './index';
import { vd } from './helper/general-helper';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  const automizer = new Automizer({
    templateDir,
    outputDir
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts');

  const dataSmaller = {
    series: [
      { label: 'series s1' },
      { label: 'series s2' }
    ],
    categories: [
      { label: 'item test r1', values: [ 10, null ] },
      { label: 'item test r2', values: [ 12, 18 ] },
    ],
  }

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setChartData(dataSmaller),
      ]);
    })
    .write(`modify-chart-stacked-bars.test.pptx`)

  vd(result)
};

run().catch((error) => {
  console.error(error);
});
