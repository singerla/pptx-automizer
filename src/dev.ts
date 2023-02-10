import { vd } from './helper/general-helper';
import Automizer, { modify } from './index';

const outputName = 'modify-chart-props.test.pptx';
const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  // You can enable 'archiveType' and set mode: 'fs'
  // This will extract all templates and output to disk.
  // It will not improve performance, but it can help debugging:
  // You don't have to manually extract pptx contents, which can
  // be annoying if you need to look inside your files.
  // archiveType: {
  //   mode: 'fs',
  //   baseDir: `${__dirname}/../__tests__/pptx-cache`,
  //   workDir: outputName,
  //   cleanupWorkDir: true,
  // },
  rootTemplate: 'RootTemplate.pptx',
  presTemplates: [`ChartBarsStacked.pptx`],
  removeExistingSlides: true,
  cleanup: true,
  compression: 0,
});

const run = async () => {
  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };
  
  const result = await automizer
    .addSlide('ChartBarsStacked.pptx', 1, (slide) => {
      slide.modifyElement('BarsStacked', [
        // This needs to be worked out:
        modify.setPlotArea({
          // Plot area width is a share of chart space.
          // We shrink it to 50% of available chart width.
          w: 0.4,
          h: 0.4,
          x: 0.0,
          y: 0.0
        }),
        // Label area position and dimensions. Automatically sets the label visible.
        modify.setLabelArea({
          w: 0.2,
          h: 1.0,
          x: 1.0,
          y: 0.0
        }),
        // Hides the label by zeroing label area
        modify.setLabelHidden(),
        // We can as well set chart data to insert our custom values.
        // modify.setChartData(dataSmaller),

        // Dump the shape:
        // modify.dump,
      ]);
    })
    .write(outputName);

  vd('It took ' + result.duration.toPrecision(2) + 's');
};

run().catch((error) => {
  console.error(error);
});
