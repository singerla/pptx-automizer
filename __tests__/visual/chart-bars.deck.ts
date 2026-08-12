/**
 * Golden deck `chart-bars` (Tier 3): setChartData growing and shrinking
 * series/categories on bar charts — mirrors modify-existing-chart and
 * modify-chart-stacked-bars, but judged by rendered pixels instead of XML.
 */
import Automizer, { ChartData, modify } from '../../src/index';
import { expectDeckMatchesBaselines } from './helpers/golden-deck';

test('golden deck: chart-bars', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../pptx-templates`,
    outputDir: `${__dirname}/../pptx-output`,
  });

  const grown: ChartData = {
    series: [
      { label: 'series s1' },
      { label: 'series s2' },
      { label: 'series s3' },
      { label: 'series s4' },
    ],
    categories: [
      { label: 'item test r1', values: [10, 16, 12, 15] },
      { label: 'item test r2', values: [12, 18, 15, 15] },
      { label: 'item test r3', values: [14, 12, 11, 15] },
      { label: 'item test r4', values: [8, 11, 9, 15] },
      { label: 'item test r5', values: [6, 15, 7, 15] },
      { label: 'item test r6', values: [16, 16, 9, 3] },
    ],
  };

  const shrunk: ChartData = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, null] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  const outputFile = 'visual-chart-bars.pptx';
  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`ChartBarsStacked.pptx`, 'bars')
    .addSlide('charts', 2, (slide) => {
      slide.modifyElement('ColumnChart', [
        modify.setChartData(<ChartData>{
          series: [
            { label: 'series 1' },
            { label: 'series 2' },
            { label: 'series 3' },
          ],
          categories: [
            { label: 'cat 2-1', values: [50, 50, 20] },
            { label: 'cat 2-2', values: [14, 50, 20] },
            { label: 'cat 2-3', values: [15, 50, 20] },
            { label: 'cat 2-4', values: [26, 50, 20] },
          ],
        }),
      ]);
    })
    .addSlide('bars', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(grown)]);
    })
    .addSlide('bars', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(shrunk)]);
    })
    .write(outputFile);

  await expectDeckMatchesBaselines('chart-bars', outputFile, 4);
});
