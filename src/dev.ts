import Automizer from './index';
import pptxgen from 'pptxgenjs';

let dataChartAreaLine = [
  {
    name: 'Actual Sales',
    labels: [
      'Jan',
      'Feb',
      'Mar',
      'Apr',
      'May',
      'Jun',
      'Jul',
      'Aug',
      'Sep',
      'Oct',
      'Nov',
      'Dec',
    ],
    values: [
      1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123,
      15121,
    ],
  },
  {
    name: 'Projected Sales',
    labels: [
      'Jan',
      'Feb',
      'Mar',
      'Apr',
      'May',
      'Jun',
      'Jul',
      'Aug',
      'Sep',
      'Oct',
      'Nov',
      'Dec',
    ],
    values: [
      1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121,
    ],
  },
];

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  let presPptxGen = new pptxgen();

  let slide = presPptxGen.addSlide();
  let textboxText = 'Hello World from PptxGenJS!';
  let textboxOpts: pptxgen.TextPropsOptions = {
    x: 1,
    y: 1,
    color: '363636',
    objectName: 'Text 1',
  };
  slide.addText(textboxText, textboxOpts);
  slide.addImage({
    path: 'https://upload.wikimedia.org/wikipedia/en/a/a9/Example.jpg',
    objectName: 'Image 1',
    x: 1,
    y: 2,
  });

  await presPptxGen.writeFile({
    fileName: templateDir + '/presPptxGenTmp.pptx',
  });

  const automizer = new Automizer({
    templateDir,
    outputDir,
    removeExistingSlides: true,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`presPptxGenTmp.pptx`, 'presPptxGenTmp')
    .load(`SlideWithShapes.pptx`, 'shapes')
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`EmptySlide.pptx`, 'emptySlide');

  // pres.addSlide('emptySlide', 1, async (slide) => {
  //   slide.generate((pptxGenJSSlide, objectName) => {
  //     pptxGenJSSlide.addText('Test', {
  //       x: 1,
  //       y: 1,
  //       color: '363636',
  //       objectName,
  //     });
  //   }, 'custom object name');
  // });
  //
  // pres.addSlide('charts', 2, (slide) => {
  //   slide.modifyElement('ColumnChart', [
  //     modify.setChartData(<ChartData>{
  //       series: [
  //         { label: 'series 1' },
  //         { label: 'series 2' },
  //         { label: 'series 3' },
  //       ],
  //       categories: [
  //         { label: 'cat 2-1', values: [50, 50, 20] },
  //         { label: 'cat 2-2', values: [14, 50, 20] },
  //         { label: 'cat 2-3', values: [15, 50, 20] },
  //         { label: 'cat 2-4', values: [26, 50, 20] },
  //       ],
  //     }),
  //   ]);
  // });
  //
  // pres.addSlide('emptySlide', 1, async (slide) => {
  //   slide.generate((pptxGenJSSlide, objectName) => {
  //     pptxGenJSSlide.addImage({
  //       path: 'https://upload.wikimedia.org/wikipedia/en/a/a9/Example.jpg',
  //       x: 1,
  //       y: 2,
  //       objectName,
  //     });
  //   });
  // });

  pres.addSlide('shapes', 1, (slide) => {
    slide.addElement('presPptxGenTmp', 1, 'Text 1');
    slide.addElement('presPptxGenTmp', 1, 'Image 1');
    pres.addSlide('emptySlide', 1, async (slide) => {
      slide.addElement('charts', 2, 'ColumnChart');

      slide.generate((pptxGenJSSlide, objectName, pptxGenJs) => {
        pptxGenJSSlide.addChart('line', dataChartAreaLine, {
          x: 1,
          y: 1,
          w: 8,
          h: 4,
          objectName,
        });
      });
      //
      // slide.generate((pptxGenJSSlide, objectName, pptxGenJs) => {
      //   pptxGenJSSlide.addChart('line', dataChartAreaLine, {
      //     x: 3,
      //     y: 1,
      //     w: 6,
      //     h: 2,
      //     objectName,
      //   });
      // }, 'MyLineChart');
    });

    pres.addSlide('emptySlide', 1, async (slide) => {
      slide.addElement('charts', 2, 'ColumnChart');

      slide.generate((pptxGenJSSlide, objectName, pptxGenJs) => {
        pptxGenJSSlide.addChart('line', dataChartAreaLine, {
          x: 1,
          y: 1,
          w: 8,
          h: 4,
          objectName,
        });
      });
      //
      // slide.generate((pptxGenJSSlide, objectName, pptxGenJs) => {
      //   pptxGenJSSlide.addChart('line', dataChartAreaLine, {
      //     x: 3,
      //     y: 1,
      //     w: 6,
      //     h: 2,
      //     objectName,
      //   });
      // }, 'MyLineChart');
    });
    pres.write(`myOutputPresentation.pptx`).then((summary) => {
      console.log(summary);
    });
  });
};

run().catch((error) => {
  console.error(error);
});
