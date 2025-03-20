import Automizer from '../src/automizer';
import { ModifyTextHelper } from '../src';

test('Append and modify slideMastes and use slideLayouts', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    verbosity: 2,
  });

  const pres = await automizer
    .loadRoot(`EmptyTemplate.pptx`)
    .load(`SlideWithNotes.pptx`, 'notes')
    .load('SlidesWithAdditionalMaster.pptx')
    .load('SlideWithShapes.pptx')
    .load('SlideWithCharts.pptx')

    // Import another slide master and all its slide layouts:
    .addMaster('SlidesWithAdditionalMaster.pptx', 1, (master) => {
      master.modifyElement(
        `MasterRectangle`,
        ModifyTextHelper.setText('my text on master'),
      );
      master.addElement(`SlideWithCharts.pptx`, 1, 'StackedBars');
    })
    .addMaster('SlidesWithAdditionalMaster.pptx', 2, (master) => {
      master.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
    })

    // Add a slide (which might require an imported master):
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // To use the original master from 'SlidesWithAdditionalMaster.pptx',
      // we can skip the argument. The required slideMaster & layout will be
      // auto imported.
      slide.useSlideLayout();
    })

    // Add a slide and use the source slideLayout:
    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // To use the original master from 'SlidesWithAdditionalMaster.pptx',
      // we can skip the argument.
      slide.useSlideLayout();
    })

    // Add a slide (which might require an imported master):
    .addSlide('notes', 1, (slide) => {
      // use another master, e.g. the imported one from 'SlidesWithAdditionalMaster.pptx'
      // You need to pass the index of the desired layout after all
      // related layouts of all imported masters have been added to rootTemplate.
      slide.useSlideLayout(26);
    })

    .write(`add-slide-master.test.pptx`);

  expect(pres.masters).toBe(3);
});
