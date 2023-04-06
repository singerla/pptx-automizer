import Automizer, { modify } from '../src/index';
import { ChartData } from '../src/types/chart-types';

const data = {
  series: [{ label: 'series 1' }, { label: 'series 2' }, { label: 'series 3' }],
  categories: [
    { label: 'cat 2-1', values: [50, 50, 20] },
    { label: 'cat 2-2', values: [14, 50, 20] },
    { label: 'cat 2-3', values: [15, 50, 20] },
    { label: 'cat 2-4', values: [26, 50, 20] },
  ],
};

const dataSmall = {
  series: [{ label: 'series 1' }],
  categories: [
    { label: 'cat 1-1', values: [50] },
    { label: 'cat 1-2', values: [14] },
  ],
};

test('Add slide with charts and modify existing chart legend.', async () => {
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
        modify.setChartData(data),
        modify.minimizeChartLegend(),
      ]);
      slide.modifyElement('PieChart', [
        modify.setChartData(dataSmall),
        modify.removeChartLegend(),
      ]);
    })
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(data),
        modify.setLegendPosition({
          w: 0.3,
        }),
      ]);
    })
    .write(`modify-chart-legend.test.pptx`);

  expect(result.charts).toBe(7);
});
