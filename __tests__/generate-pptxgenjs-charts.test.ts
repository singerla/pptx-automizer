import Automizer from '../src/automizer';
import { ChartData, modify } from '../src';

const dataChartAreaLine = [
  {
    name: 'Actual Sales',
    labels: ['Jan', 'Feb', 'Mar'],
    values: [1500, 4600, 5156],
  },
  {
    name: 'Projected Sales',
    labels: ['Jan', 'Feb', 'Mar'],
    values: [1000, 2600, 3456],
  },
];

test('generate a chart with pptxgenjs and add it to a template slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithCharts.pptx`, 'charts');

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add generated contents from scratch:
    slide.generate((pSlide, pptxGenJs) => {
      pSlide.addChart(pptxGenJs.ChartType.line, dataChartAreaLine, {
        x: 1,
        y: 1,
        w: 8,
        h: 4,
      });
    });
  });

  // Mix the created chart with modified existing chart
  pres.addSlide('charts', 2, (slide) => {
    slide.modifyElement('ColumnChart', [
      modify.setChartData(<ChartData>{
        series: [{ label: 'series 1' }, { label: 'series 3' }],
        categories: [
          { label: 'cat 2-1', values: [50, 50] },
          { label: 'cat 2-2', values: [14, 50] },
        ],
      }),
    ]);
  });

  const result = await pres.write(`generate-pptxgenjs-charts.test.pptx`);

  expect(result.charts).toBe(4);
});
