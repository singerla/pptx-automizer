import Automizer, { ChartData, modify, TableRow, TableRowStyle } from './index';
import { vd } from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
});

const run = async () => {
  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartWaterfall.pptx`, 'charts');

  const result = await pres
    .addSlide('charts', 1, (slide) => {
      slide.addElement('charts', 1, 'Waterfall 1', [
        modify.setChartData(<ChartData>{
          series: [{ label: 'series 1' }],
          categories: [
            { label: 'cat 2-1', values: [50] },
            { label: 'cat 2-2', values: [14] },
            { label: 'cat 2-3', values: [15] },
            { label: 'cat 2-4', values: [26] },
          ],
        }),
      ]);
    })
    .write(`modify-existing-waterfall-chart.test.pptx`);
};

run().catch((error) => {
  console.error(error);
});
