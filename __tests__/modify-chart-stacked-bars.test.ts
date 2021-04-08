import Automizer, { modify } from '../src/index';

test('create presentation, add vertical lines chart, add a series, add categories.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`ChartBarsStacked.pptx`, 'charts');

  const data = {
    series: [
      { label: 'series s1' }, 
      { label: 'series s2' },
      { label: 'series s3' },
      { label: 'series s4' },
    ],
    categories: [
      { label: 'item test r1', values: [ 10, 16, 12, 15 ] },
      { label: 'item test r2', values: [ 12, 18, 15, 15 ] },
      { label: 'item test r3', values: [ 14, 12, 11, 15 ] },
      { label: 'item test r4', values: [ 8, 11, 9, 15 ] },
      { label: 'item test r5', values: [ 6, 15, 7, 15 ] },
      { label: 'item test r6', values: [ 16, 16, 9, 3 ] },
    ],
  }

  const dataSmaller = {
    series: [
      { label: 'series s1' }, 
      { label: 'series s2' }
    ],
    categories: [
      { label: 'item test r1', values: [ 10, 16 ] },
      { label: 'item test r2', values: [ 12, 18 ] },
    ],
  }

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setChartData(data),
      ]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setChartData(dataSmaller),
      ]);
    })
    .write(`modify-chart-stacked-bars.test.pptx`)

  expect(result.charts).toBe(4);
});
