import { vd } from './helper/general-helper';
import Automizer, { modify } from './index';

const automizer = (outputName) =>
  new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
    // You can enable 'archiveType' and set mode: 'fs'
    // This will extract all templates and output to disk.
    // It will not improve performance, but it can help debugging:
    // You don't have to manually extract pptx contents, which can
    // be annoying if you need to look inside your files.
    archiveType: {
      mode: 'fs',
      baseDir: `${__dirname}/../__tests__/pptx-cache`,
      workDir: outputName,
      // cleanupWorkDir: true,
    },
    rootTemplate: 'RootTemplate.pptx',
    presTemplates: [`SlidesWithAdditionalMaster.pptx`],
    removeExistingSlides: false,
    cleanup: false,
    compression: 0,
  });

const run = async () => {
  // const outputName1 = 'import-template-master.test.pptx';
  // const result = await automizer(outputName1)
  //   // These two will be ok, because `SlidesWithAdditionalMaster.pptx` and
  //   // 'RootTemplate.pptx' have at least one slide master.
  //   // BUT: Master #1 of 'RootTemplate.pptx' will be used for the inserted
  //   // slides, which is probably not intended.
  //   .addSlide('SlidesWithAdditionalMaster.pptx', 1, (slide) => {})
  //   .addSlide('SlidesWithAdditionalMaster.pptx', 2, (slide) => {})
  //
  //   // Adding slide 3 will result to a corrupted pptx, because it is related to
  //   // a second slide master. As long as there is only one slide master present
  //   // in RootTemplate.pptx, slide 3 will be related to a missing master:
  //   .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {})
  //   .write(outputName1);
  // vd('It took ' + result.duration.toPrecision(2) + 's');

  // It should be like this:
  const outputName2 = 'import-template-master-2.test.pptx';
  const result2 = await automizer(outputName2)
    // Import another slide master and all its slide layouts:
    .addMaster('SlidesWithAdditionalMaster.pptx', 1)
    .addMaster('SlidesWithAdditionalMaster.pptx', 2)

    // Add a slide (which might require an imported master):
    // .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
    //   // use another master, e.g. the imported one from 'SlidesWithAdditionalMaster.pptx'
    //   slide.useMaster('myMaster#2');
    // })
    // Add a slide and use another master:
    // .addSlide('SlidesWithAdditionalMaster.pptx', 2, (slide) => {
    //   // find an imported master/layout by name
    //   slide.useMaster('Orange Design', 'Leer');
    // })
    .write(outputName2);

  vd('It took ' + result2.duration.toPrecision(2) + 's');
};

run().catch((error) => {
  console.error(error);
});
