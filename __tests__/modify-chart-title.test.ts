import Automizer, { modify } from '../src/index';

test('modify chart title.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts');


  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        modify.setChartTitle('New Title'),
      ]);
    })
    .write(`modify-chart-title.test.pptx`);

  expect(result.charts).toBe(2);
});
