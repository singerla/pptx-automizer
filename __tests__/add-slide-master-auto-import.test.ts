import Automizer from '../src/automizer';
import { ModifyTextHelper } from '../src';

test('Auto-import source slideLayout and -master', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    autoImportSlideMasters: true,
  });

  const pres = await automizer
    .loadRoot(`EmptyTemplate.pptx`)
    .load('SlidesWithAdditionalMaster.pptx')
    .load('SlideMasterBackgrounds.pptx')
    .load('SlideMasters.pptx')

    // We can disable .addMaster according to "autoImportSlideMasters: true"
    // .addMaster('SlidesWithAdditionalMaster.pptx', 1)

    .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
      // We can also disable "slide.useSlideLayout()"
      // with "autoImportSlideMasters: true"
      // slide.useSlideLayout();
    })
    .addSlide('SlideMasters.pptx', 1)
    .addSlide('SlidesWithAdditionalMaster.pptx', 1)
    .addSlide('SlideMasters.pptx', 3)
    .addSlide('SlideMasters.pptx', 2)
    .addSlide('SlideMasterBackgrounds.pptx', 2)
    .write(`add-slide-master-auto-import.test.pptx`);

  expect(pres.masters).toBe(6);
});
