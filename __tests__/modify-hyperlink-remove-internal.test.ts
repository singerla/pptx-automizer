import Automizer, { modify } from '../src/index';
import { vd } from '../src/helper/general-helper';

test('delete internal hyperlink - using removeHyperlink helper', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    verbosity: 0
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `delete-hyperlink-internal.test.pptx`;

  const result = await pres
    .addSlide('link', 1, (slide) => {
      slide.modifyElement('LinkToSlide', [modify.removeHyperlink()]);
    })
    .write(outputFile);

  console.log(result);
});
