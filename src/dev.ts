import Automizer, { ChartData, modify } from './index';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    removeExistingSlides: true,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`EmptySlide.pptx`, 'emptySlide')

  pres.addSlide('emptySlide', 1, async (slide) => {
    slide.addElement('chart', 3, '33174534-89bf-4326-8085-4b6938d36f7d');
    // slide.addElement('charts', 2, 'ColumnChart');

    // slide.generate((pptxGenJSSlide, objectName, pptxGenJs) => {
    //   let dataChartAreaLine = [
    //     {
    //       name: 'Actual Sales',
    //       labels: [
    //         'Jan',
    //         'Feb',
    //         'Mar',
    //         'Apr',
    //         'May',
    //         'Jun',
    //         'Jul',
    //         'Aug',
    //         'Sep',
    //         'Oct',
    //         'Nov',
    //         'Dec',
    //       ],
    //       values: [
    //         1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123,
    //         15121,
    //       ],
    //     },
    //     {
    //       name: 'Projected Sales',
    //       labels: [
    //         'Jan',
    //         'Feb',
    //         'Mar',
    //         'Apr',
    //         'May',
    //         'Jun',
    //         'Jul',
    //         'Aug',
    //         'Sep',
    //         'Oct',
    //         'Nov',
    //         'Dec',
    //       ],
    //       values: [
    //         1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123,
    //         12121,
    //       ],
    //     },
    //   ];
    //
    //   pptxGenJSSlide.addChart('line', dataChartAreaLine, {
    //     x: 1,
    //     y: 1,
    //     w: 8,
    //     h: 4,
    //     objectName,
    //   });
    // });
  });

  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
