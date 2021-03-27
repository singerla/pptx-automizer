import Automizer, { modify } from '../src/index';
import { ChartData } from '../src/types/chart-types';

test('create presentation, add slide with charts from template and modify existing chart.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts');

  const result = await pres
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            {label: 'series 1'},
            {label: 'series 2'},
            {label: 'series 3'},
          ],
          categories: [
            {label: 'cat 2-1', values: [50, 50, 20]},
            {label: 'cat 2-2', values: [14, 50, 20]},
            {label: 'cat 2-3', values: [15, 50, 20]},
            {label: 'cat 2-4', values: [26, 50, 20]}
          ]
        })
      ]);
    })
    .write(`modify-existing-chart.test.pptx`);

  expect(result.charts).toBe(3);
});
