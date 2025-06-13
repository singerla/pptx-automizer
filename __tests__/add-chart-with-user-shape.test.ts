import Automizer, { modify } from '../src/index';

test('add charts with external data and mix with default charts', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts')
    .load(`ChartUserShape.pptx`, 'chartUserShape');

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.addElement('chartUserShape', 1, 'BarsStackedWithUserShape');
    });

  const result = await pres.write(`add-chart-with-user-shape.test.pptx`);

  expect(result.charts).toBe(4);
});
