import Automizer, { modify } from '../src/index';
import { ChartData } from '../src/types/chart-types';

test('modify chart axis.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartAxis.pptx`, 'charts');

  const dataScatter = <ChartData>(<unknown>{
    series: [
      { label: 'series s1' },
      { label: 'series s2' },
      { label: 'series s3' },
    ],
    categories: [
      {
        label: 'r1',
        values: [
          { x: 10, y: 20 },
          { x: 9, y: 30 },
          { x: 19, y: 40 },
        ],
      },
      {
        label: 'r2',
        values: [
          { x: 21, y: 11 },
          { x: 8, y: 31 },
          { x: 18, y: 41 },
        ],
      },
      {
        label: 'r3',
        values: [
          { x: 22, y: 28 },
          { x: 7, y: 26 },
          { x: 17, y: 36 },
        ],
      },
      {
        label: 'r4',
        values: [
          { x: 13, y: 13 },
          { x: 16, y: 28 },
          { x: 26, y: 38 },
        ],
      },
      {
        label: 'r5',
        values: [
          { x: 18, y: 24 },
          { x: 15, y: 24 },
          { x: 25, y: 34 },
        ],
      },
      {
        label: 'r6',
        values: [
          { x: 28, y: 34 },
          { x: 25, y: 34 },
          { x: 35, y: 44 },
        ],
      },
    ],
  });

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('Scatter', [
        modify.setChartScatter(dataScatter),

        /**
         * Please notice: It will only work if the value to update is not set to
         * "Auto" in powerpoint. Only manually scaled min/max can be altered by this.
         *
         * min: 10 will not take any effect, while max: 30 does.
         */
        modify.setAxisRange({
          axisIndex: 0,
          min: 10,
          max: 30,
        }),

        modify.setAxisRange({
          axisIndex: 1,
          min: 1,
          max: 100,
        }),
      ]);
    })
    .write(`modify-chart-axis.test.pptx`);

  expect(result.charts).toBe(2);
});
