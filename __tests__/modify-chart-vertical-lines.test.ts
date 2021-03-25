import Automizer, { modify } from '../src/index';

test('create presentation, add and modify a vertical lines chart.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`ChartLinesVertical.pptx`, 'charts');

  const data3 = {
    series: [
      { label: 'series s1' }, 
      { label: 'series s2' },
      { label: 'series s3' }
    ],
    categories: [
      { label: 'item test r1', yValue: 10, values: [10, 45, 5] },
      { label: 'item test r2', yValue: 9, values: [20, 35, 6] },
      { label: 'item test r3', yValue: 8, values: [15, 25, 7] },
      { label: 'item test r4', yValue: 7, values: [12, 15, 8] },
      { label: 'item test r5', yValue: 6, values: [8, 8, 9] },
      { label: 'item test r6', yValue: 5, values: [9, 7, 10] },
      { label: 'item test r7', yValue: 4, values: [6, 6, 11] },
      { label: 'item test r8', yValue: 3, values: [11, 5, 12] },
      { label: 'item test r9', yValue: 2, values: [12, 4, 13] },
      { label: 'item test r10', yValue: 1, values: [19, 3, 14] },
    ],
  }

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('DotMatrix', [
        modify.setChartVerticalLines(data3),
      ]);
    })
    .write(`modify-chart-vertical-lines.test.pptx`)

  expect(result.charts).toBe(2);
});
