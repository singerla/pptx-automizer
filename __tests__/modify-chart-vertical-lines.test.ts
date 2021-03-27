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
      { label: 'item test r1', y: 10, values: [ 10, 16, 12 ] },
      { label: 'item test r2', y: 9, values: [ 12, 18, 15 ] },
      { label: 'item test r3', y: 8, values: [ 14, 12, 11 ] },
      { label: 'item test r4', y: 7, values: [ 8, 11, 9 ] },
      { label: 'item test r5', y: 6, values: [ 6, 15,7 ] },
      { label: 'item test r6', y: 5, values: [ 16, 16, 9 ] },
      { label: 'item test r7', y: 4, values: [ 10, 13, 12 ] },
      { label: 'item test r8', y: 3, values: [ 11, 12, 14 ] },
      { label: 'item test r9', y: 2, values: [ 9, 7, 11 ] },
      { label: 'item test r10', y: 1, values: [ 7, 5, 17 ] }
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
