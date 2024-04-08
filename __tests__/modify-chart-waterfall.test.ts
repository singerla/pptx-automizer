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

test('Add slide with charts and modify waterfall data and total column.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartWaterfall.pptx`, 'charts');

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('Waterfall 1', [
        modify.setWaterFallColumnTotalToLast(), //set total to last
      ]);
    })
    .addSlide('charts', 1, (slide) => { //Modify Data
      slide.modifyElement('Waterfall 1', [
      modify.setExtendedChartData({
        series: [{ label: 'series 1' }],
        categories: [
          { label: 'cat 2-1', values: [100] },
          { label: 'cat 2-3', values: [50] },
          { label: 'cat 2-4', values: [-40] },
        ],
      }),
    ]);
  })
    .addSlide('charts', 1, (slide) => { //Modify Data and set Total to middle
      slide.modifyElement('Waterfall 1', [
      modify.setExtendedChartData({
        series: [{ label: 'series 1' }],
        categories: [
          { label: 'cat 2-1', values: [100] },
          { label: 'cat 2-3', values: [50] },
          { label: 'sub total', values: [150] },
          { label: 'cat 2-4', values: [-40] }
        ],
      }),
      modify.setWaterFallColumnTotalToLast(2)
    ]);
  })
    .addSlide('charts', 1, (slide) => { //Modify Data and set Total to last
      slide.modifyElement('Waterfall 1', [
      modify.setExtendedChartData({
        series: [{ label: 'series 1' }],
        categories: [
          { label: 'cat 2-1', values: [100] },
          { label: 'cat 2-3', values: [50] },
          { label: 'sub total', values: [150] },
          { label: 'cat 2-4', values: [-40] },
          { label: 'total', values: [330] },
        ],
      }),
      modify.setWaterFallColumnTotalToLast()
    ]);
  })
    .write(`modify-chart-waterfall.test.pptx`);

  expect(result.charts).toBe(8);
});
