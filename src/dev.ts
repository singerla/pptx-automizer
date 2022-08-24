import Automizer, { ChartData, modify, TableRow, TableRowStyle } from './index';
import { vd } from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
});

const run = async () => {
  const pres = automizer
    .loadRoot(`ChartLinesVerticalImageMarkers.pptx`)
    .load(`ChartLinesVerticalImageMarkers.pptx`, 'charts');

  const data = {
    series: [
      { label: 'series s1' },
      { label: 'series s2' },
      { label: 'series s3' },
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
      slide.modifyElement('DotMatrixImageMarkers', [
        modify.setChartVerticalLines(dataSmaller),
      ]);
    })
    .write(`modify-chart-vertical-lines-marker.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
