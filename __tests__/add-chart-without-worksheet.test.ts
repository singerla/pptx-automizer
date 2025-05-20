import Automizer, { modify } from '../src/index';

test('add charts with external data and mix with default charts', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts')
    .load(`ChartWithExternalData.pptx`, 'externalData');

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  pres
    .addSlide('externalData', 1)
    .addSlide('externalData', 1, (slide) => {
      slide.removeElement('Gráfico 7');
      slide.addElement('charts', 1, 'BarsStacked');
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.removeElement('BarsStacked');
      // Add chart with external data to another slide
      slide.addElement('externalData', 1, 'Gráfico 7');
    });

  const result = await pres.write(`add-chart-without-worksheet.test.pptx`);

  expect(result.charts).toBe(9);
});
