import Automizer from '../src/automizer';
import { ModifyShapeHelper, ModifyTextHelper } from '../src';

test('Import layout from another template and merge with auto-mapping placeholders', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    autoImportSlideMasters: false,
    cleanupPlaceholders: false,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlidesWithAdditionalMaster.pptx`)
    .load(`EmptySlidePlaceholders.pptx`)
    .addMaster(`SlidesWithAdditionalMaster.pptx`, 2);

  pres.addSlide('EmptySlidePlaceholders.pptx', 2, async (slide) => {
    slide.mergeIntoSlideLayout('Titel und Inhalt');

    slide.addElement(
      'EmptySlidePlaceholders.pptx',
      2,
      '@TitleNoPlaceholder',
      ModifyTextHelper.setText('Test orig add'),
    );
    slide.modifyElement('@TitleNoPlaceholder', [
      ModifyShapeHelper.roundedCorners(1400),
    ]);
    slide.modifyElement('@TitleNoPlaceholder', [
      ModifyTextHelper.setText('Test 1235'),
    ]);
  });

  await pres.write(`add-slide-master-merge-layout.test.pptx`);
});
