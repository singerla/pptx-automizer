import { vd } from './helper/general-helper';
import Automizer, { modify } from './index';

const outputName = 'create-presentation-file-proxy.test.pptx';
const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  // Streaming is only implemented for jszip
  // archiveType: {
  //   mode: 'fs',
  //   baseDir: `${__dirname}/../__tests__/pptx-cache`,
  //   workDir: outputName,
  //   cleanupWorkDir: true,
  // },
  rootTemplate: 'RootTemplateWithImages.pptx',
  presTemplates: [
    `RootTemplate.pptx`,
    `SlideWithImages.pptx`,
    `ChartBarsStacked.pptx`,
  ],
  removeExistingSlides: true,
  cleanup: true,
  compression: 1,
});

const run = async () => {
  // const pres = automizer
  //   .loadRoot(`RootTemplateWithImages.pptx`)
  //   .load(`RootTemplate.pptx`, 'root')
  //   .load(`SlideWithImages.pptx`, 'images')
  //   .load(`ChartBarsStacked.pptx`, 'charts');

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  const stream = await automizer
    .addSlide('ChartBarsStacked.pptx', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
      slide.addElement('ChartBarsStacked.pptx', 1, 'BarsStacked', [
        modify.setChartData(dataSmaller),
      ]);
    })
    .addSlide('SlideWithImages.pptx', 1)
    .addSlide('RootTemplate.pptx', 1, (slide) => {
      slide.addElement('ChartBarsStacked.pptx', 1, 'BarsStacked', [
        modify.setChartData(dataSmaller),
      ]);
    })
    .addSlide('ChartBarsStacked.pptx', 1, (slide) => {
      slide.addElement('SlideWithImages.pptx', 2, 'imageJPG');
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .addSlide('ChartBarsStacked.pptx', 1, (slide) => {
      slide.addElement('SlideWithImages.pptx', 2, 'imageJPG');
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .stream({
      compressionOptions: {
        level: 2,
      },
    });

  // vd(stream);

  // stream.pipe(process.stdout);

  // vd(pres.rootTemplate.content);
};

run().catch((error) => {
  console.error(error);
});
