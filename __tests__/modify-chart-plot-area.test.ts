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

test('Add slide with charts and modify chart plot area coordinates.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts');

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('StackedBars', [
        modify.setChartData(data),
        modify.setPlotArea({
          w: 0.5,
          h: 0.5,
        }),
      ]);
    })
    .write(`modify-chart-plot-area.test.pptx`);

  expect(result.charts).toBe(2);
});
