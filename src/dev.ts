import { vd } from './helper/general-helper';
import Automizer, { modify, ModifyTextHelper, XmlHelper } from './index';
import { XmlElement } from './types/xml-types';

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
    presTemplates: [
      `SlidesWithAdditionalMaster.pptx`,
      `SlideWithShapes.pptx`,
      `SlideWithCharts.pptx`,
    ],
    removeExistingSlides: true,
    cleanup: false,
    compression: 0,
  });

const run = async () => {
  const outputName2 = 'import-template-master-2.test.pptx';
  const result2 = await automizer(outputName2)
    // We can skip master import and rely on autoImportSourceSlideMaster during
    // when calling slide.useSlideLayout().

    // Import another slide master and all its slide layouts:
    // .addMaster('SlidesWithAdditionalMaster.pptx', 1, (master) => {
    //   master.modifyElement(
    //     `MasterRectangle`,
    //     ModifyTextHelper.setText('my text on master'),
    //   );
    //   master.addElement(`SlideWithCharts.pptx`, 1, 'StackedBars');
    // })
    // .addMaster('SlidesWithAdditionalMaster.pptx', 2, (master) => {
    //   // master.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
    // })

    // Add a slide (which might require an imported master):
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // To use the original master from 'SlidesWithAdditionalMaster.pptx',
      // we can skip the argument. The required slideMaster & layout will be
      // auto imported.
      slide.useSlideLayout();
    })

    // Add a slide (which might require an imported master):
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // use another master, e.g. the imported one from 'SlidesWithAdditionalMaster.pptx'
      // You need to pass the index of the desired layout after all
      // related layouts of all imported masters have been added to rootTemplate.
      slide.useSlideLayout(22);
    })

    // Add a slide (which might require an imported master):
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // To use the original master from 'SlidesWithAdditionalMaster.pptx',
      // we can skip the argument.
      slide.useSlideLayout();
    })

    .write(outputName2);

  vd('It took ' + result2.duration.toPrecision(2) + 's');
};

run().catch((error) => {
  console.error(error);
});
