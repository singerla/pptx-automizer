import Automizer, { modify } from '../src/index';
import { ChartData } from '../dist';

test('create presentation, add and modify a scatter chart with embedded point images.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    // TODO: cleanup unused marker images
    // (which need to be tracked first)
    // cleanup: true,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartScatter.pptx`, 'charts');

  const dataScatter = <ChartData>(<unknown>{
    series: [{ label: 'series s1' }],
    categories: [
      { label: 'r1', values: [{ x: 10, y: 20 }] },
      { label: 'r2', values: [{ x: 21, y: 11 }] },
      { label: 'r3', values: [{ x: 22, y: 28 }] },
      { label: 'r4', values: [{ x: 13, y: 13 }] },
    ],
  });

  const result = await pres
    .addSlide('charts', 3, (slide) => {
      slide.modifyElement('ScatterPointImages', [
        modify.setChartScatter(dataScatter),
      ]);
    })
    .write(`modify-chart-scatter-images.test.pptx`);

  // expect(result.charts).toBe(2);
});
