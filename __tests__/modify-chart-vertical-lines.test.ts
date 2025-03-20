import Automizer, { ChartData, modify } from '../src/index';

test('create presentation, add and modify a vertical lines chart.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartLinesVertical.pptx`, 'charts');

  const data: ChartData = {
    series: [
      { label: 'series s1' },
      {
        label: 'series s2',
        style: {
          color: {
            type: 'schemeClr',
            value: 'accent1',
          },
          // Disable labels for this series
          label: {
            showVal: false,
          },
        },
      },
      {
        label: 'series s3',
        style: {
          color: {
            type: 'schemeClr',
            value: 'accent1',
          },
          // All labels of a series can be styled
          // Notice: first, activate Datalabels in your template chart
          label: {
            color: {
              type: 'schemeClr',
              value: 'accent2',
            },
            isBold: false,
            size: 2200,
            showVal: true,
            showLegendKey: true,
            solidFill: {
              type: 'schemeClr',
              value: 'accent1',
            },
          },
        },
      },
    ],
    categories: [
      { label: 'item test r1', y: 10, values: [10, 16, 12] },
      { label: 'item test r2', y: 9, values: [12, 18, 15] },
      { label: 'item test r3', y: 8, values: [14, 12, 11] },
      { label: 'item test r4', y: 7, values: [8, 11, 9] },
      { label: 'item test r5', y: 6, values: [6, 15, 7] },
      { label: 'item test r6', y: 5, values: [16, 16, 9] },
      { label: 'item test r7', y: 4, values: [10, 13, 12] },
      { label: 'item test r8', y: 3, values: [11, 12, 14] },
      { label: 'item test r9', y: 2, values: [9, 7, 11] },
      { label: 'item test r10', y: 1, values: [7, 5, 17] },
    ],
  };

  const dataSmaller = {
    series: [{ label: 'series s1' }],
    categories: [
      { label: 'item test r1', y: 10, values: [10] },
      { label: 'item test r2', y: 9, values: [12] },
      { label: 'item test r3', y: 8, values: [14] },
    ],
  };

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('DotMatrix', [modify.setChartVerticalLines(data)]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('DotMatrix', [
        modify.setChartVerticalLines(dataSmaller),
      ]);
    })
    .write(`modify-chart-vertical-lines.test.pptx`);

  expect(result.charts).toBe(4);
});
