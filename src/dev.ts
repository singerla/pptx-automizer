import Automizer, {
  ChartData,
  modify,
  TableRow,
  TableRowStyle,
  XmlHelper,
} from './index';
import { vd } from './helper/general-helper';
import ModifyPresentationHelper from './helper/modify-presentation-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
});

const run = async () => {
  // const pres = automizer
  //   .loadRoot(`RootTemplate.pptx`)
  //   .load(`SlideWithCharts.pptx`, 'charts');
  //
  // const result = await pres
  //   .addSlide('charts', 2, (slide) => {
  //     slide.modifyElement('ColumnChart', [
  //       modify.setChartData(<ChartData>{
  //         series: [
  //           { label: 'series 1' },
  //           { label: 'series 2' },
  //           { label: 'series 3' },
  //         ],
  //         categories: [
  //           { label: 'cat 2-1', values: [50, 50, 20] },
  //           { label: 'cat 2-2', values: [14, 50, 20] },
  //           { label: 'cat 2-3', values: [15, 50, 20] },
  //           { label: 'cat 2-4', values: [26, 50, 20] },
  //         ],
  //       }),
  //     ]);
  //   })
  //   .write(`modify-existing-chart.test.pptx`);

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'EmptySlide')
    .load(`ChartWaterfall.pptx`, 'ChartWaterfall')
    .load(`ChartBarsStacked.pptx`, 'ChartBarsStacked');

  const result = await pres
    .addSlide('EmptySlide', 1, (slide) => {
      // slide.addElement('ChartBarsStacked', 1, 'BarsStacked', [
      //   modify.setChartData(<ChartData>{
      //     series: [{ label: 'series 1' }],
      //     categories: [
      //       { label: 'cat 2-1', values: [50] },
      //       { label: 'cat 2-2', values: [14] },
      //       { label: 'cat 2-3', values: [15] },
      //       { label: 'cat 2-4', values: [26] },
      //     ],
      //   }),
      // ]);

      slide.addElement('ChartWaterfall', 1, 'Waterfall 1', [
        modify.setExtendedChartData(<ChartData>{
          series: [{ label: 'series 1' }],
          categories: [
            { label: 'cat 2-1', values: [50] },
            { label: 'cat 2-2', values: [14] },
            { label: 'cat 2-3', values: [15] },
            { label: 'cat 2-4', values: [26] },
            { label: 'cat 2-4', values: [26] },
            { label: 'cat 2-4', values: [26] },
            { label: 'cat 2-4', values: [26] },
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
