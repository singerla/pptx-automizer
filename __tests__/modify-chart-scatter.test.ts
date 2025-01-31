import Automizer, {
  ChartValueStyle,
  LabelPosition,
  modify,
} from '../src/index';
import { ChartData } from '../dist';

test('create presentation, add and modify a scatter chart.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartScatter.pptx`, 'charts');

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

  const dataScatterSmall = <ChartData>(<unknown>{
    series: [{ label: 'series s1' }],
    categories: [
      {
        label: 'r1',
        values: [{ x: 10, y: 20 }],
        styles: [
          {
            color: {
              type: 'srgbClr',
              value: 'cccccc',
            },
            label: {
              size: 1600,
            },
          },
        ],
      },
      {
        label: 'r2',
        values: [{ x: 21, y: 11 }],
        styles: [
          {
            color: {
              type: 'srgbClr',
              value: 'cc00cc',
            },
            label: {
              size: 1800,
            },
          },
        ],
      },
      {
        label: 'r3',
        values: [{ x: 31, y: 21 }],
        styles: <ChartValueStyle[]>[
          {
            color: {
              type: 'srgbClr',
              value: 'cccc00',
            },
            label: {
              size: 2000,
              showVal: false,
            },
          },
        ],
      },
    ],
  });

  const result = await pres
    .addSlide('charts', 4, (slide) => {
      slide.modifyElement('ScatterPointLabel', [
        modify.setChartScatter(dataScatterSmall),
      ]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('Scatter', [modify.setChartScatter(dataScatter)]);
    })
    .addSlide('charts', 2, (slide) => {
      dataScatter.categories[0].styles = [
        {
          color: {
            type: 'srgbClr',
            value: 'cccccc',
          },
        },
      ];

      slide.modifyElement('ScatterPoint', [
        modify.setChartScatter(dataScatter),
      ]);
    })
    .write(`modify-chart-scatter.test.pptx`);

  expect(result.charts).toBe(6);
});
