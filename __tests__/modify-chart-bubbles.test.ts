import Automizer, { modify } from '../src/index';
import {ChartData} from '../dist';

test('create presentation, add and modify a bubble chart.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`ChartBubbles.pptx`, 'charts');

  const dataBubbles = <ChartData><unknown>{
    series: [
      {label: 'series s1'},
      {label: 'series s2'},
      {label: 'series s3'}
    ],
    categories: [
      {label: 'r1', values: [{x: 10, y: 20, size: 1}, {x: 9, y: 30, size: 5}, {x: 19, y: 40, size: 1}]},
      {label: 'r2', values: [{x: 21, y: 11, size: 4}, {x: 8, y: 31, size: 6}, {x: 18, y: 41, size: 2}]},
      {label: 'r3', values: [{x: 22, y: 28, size: 3}, {x: 7, y: 26, size: 5}, {x: 17, y: 36, size: 1}]},
      {label: 'r4', values: [{x: 13, y: 13, size: 2}, {x: 16, y: 28, size: 6}, {x: 26, y: 38, size: 2}]},
      {label: 'r5', values: [{x: 18, y: 24, size: 3}, {x: 15, y: 24, size: 4}, {x: 25, y: 34, size: 2}]},
      {label: 'r6', values: [{x: 28, y: 34, size: 1}, {x: 25, y: 34, size: 1}, {x: 35, y: 44, size: 1}]},
    ],
  }

  const dataBubblesSmaller = <ChartData><unknown>{
    series: [
      { label: 'series s1' },
    ],
    categories: [
      { label: 'r1', values: [ {x: 10, y: 20, size: 1} ]},
      { label: 'r2', values: [ {x: 21, y: 11, size: 4} ]},
    ],
  }

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('Bubbles', [
        modify.setChartBubbles(dataBubbles),
      ]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('Bubbles', [
        modify.setChartBubbles(dataBubblesSmaller),
      ]);
    })
    .write(`modify-chart-bubbles.test.pptx`)

  expect(result.charts).toBe(4);
});
